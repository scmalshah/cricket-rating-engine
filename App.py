import streamlit as st
import pandas as pd
import numpy as np
import os
import json
import re
import plotly.graph_objects as go
import plotly.express as px

st.set_page_config(page_title="BPL Cricket Rating Engine", layout="wide", page_icon="🏏")

# --- FILE PATHS & SECRETS ---
DATA_FILE = "current_stats.xlsx"
MAPPING_FILE = "name_mapping.json"
DRAFT_FILE = "draft_state.json"
RATINGS_FILE = "human_ratings.json"
USERS_FILE = "authorized_users.json"
WEIGHTS_FILE = "algo_weights.json"
LEADERSHIP_FILE = "leadership.json"
POOL_FILE = "pool_status.json"

ADMIN_PASSWORD = "bpladmin" 

DEFAULT_MAPPINGS = {
    "Praveenraj Starkey": "Praveenraj Starkey (Merged)",
    "Praveen Starkey": "Praveenraj Starkey (Merged)",
    "Godwin .": "Godwin (Merged)",
    "Godwin M": "Godwin (Merged)"
}

DEFAULT_WEIGHTS = {
    "bat_runs": 20.0, "bat_avg": 30.0, "bat_sr": 25.0, "bat_bpd": 10.0, "bat_bound": 15.0,
    "bowl_wkts": 30.0, "bowl_econ": 25.0, "bowl_avg": 20.0, "bowl_sr": 15.0, "bowl_extras": 10.0,
    "wt_batter_bat": 85.0, "wt_batter_field": 15.0,
    "wt_bowler_bowl": 85.0, "wt_bowler_field": 15.0,
    "wt_ar_bat": 42.5, "wt_ar_bowl": 42.5, "wt_ar_field": 15.0, "ar_multiplier": 1.3,
    "squad_size": 11, "team_budget": 240.0
}

# --- HELPER FUNCTIONS ---
def clean_col_name(c): return str(c).replace('\xa0', '').replace("'", "").strip()
def clean_prefix(name):
    if not isinstance(name, str): return ""
    name = re.sub(' +', ' ', name.replace('\xa0', ' ').strip())
    if len(name) > 2 and name[:2].isupper() and name[2:4].upper() == name[:2]:
        return name[2:].strip()
    return name
def overs_to_balls(overs):
    if pd.isna(overs): return 0
    try: return int(float(overs)) * 6 + round((float(overs) - int(float(overs))) * 10)
    except: return 0
def format_overs(balls): return float(f"{int(balls) // 6}.{int(balls) % 6}")

def load_json(filepath, default_val):
    if os.path.exists(filepath):
        try:
            with open(filepath, "r") as f: return json.load(f)
        except: return default_val
    return default_val

def save_json(filepath, data):
    with open(filepath, "w") as f: json.dump(data, f, indent=2)

def update_cap_settings():
    if "cap_squad" in st.session_state and "cap_budget" in st.session_state:
        cw = load_json(WEIGHTS_FILE, DEFAULT_WEIGHTS)
        cw["squad_size"] = st.session_state.cap_squad
        cw["team_budget"] = st.session_state.cap_budget
        save_json(WEIGHTS_FILE, cw)

def extract_name(val):
    v_str = str(val).strip()
    if not v_str or pd.isna(val) or v_str == "None" or v_str == "--- CLEAR PICK ---": 
        return ""
    v_str = v_str.replace("🔒 ", "").replace("🔒", "")
    if " (" in v_str and v_str.endswith(")"):
        return v_str.rsplit(" (", 1)[0].strip()
    return v_str.strip()

# --- LOAD STATES ---
saved_mapping = load_json(MAPPING_FILE, {})
mapping_changed = False
for original, merged in DEFAULT_MAPPINGS.items():
    if original not in saved_mapping:
        saved_mapping[original] = merged
        mapping_changed = True
if mapping_changed: save_json(MAPPING_FILE, saved_mapping)

draft_state = load_json(DRAFT_FILE, {}) 
leadership_state = load_json(LEADERSHIP_FILE, {})
pool_state = load_json(POOL_FILE, {})
human_ratings = load_json(RATINGS_FILE, {}) 
auth_users = load_json(USERS_FILE, ["Admin", "Captain 1", "Captain 2"])
TEAMS = ["Available", "Team 1", "Team 2", "Team 3", "Team 4", "Team 5"]

algo_weights = load_json(WEIGHTS_FILE, DEFAULT_WEIGHTS)
for k, v in DEFAULT_WEIGHTS.items():
    if k not in algo_weights: algo_weights[k] = v

# --- HEADER ---
st.title("🏏 BPL Cricket Rating Engine")
st.markdown("Advanced AI Rating, Live Roster Management, and Committee Scouting.")

# --- DATA PROCESSING ENGINE ---
@st.cache_data(show_spinner="Reading Excel File (Only happens once)...")
def get_raw_excel_data(file_mod_time):
    if not os.path.exists(DATA_FILE): return pd.DataFrame(), pd.DataFrame(), pd.DataFrame(), []
    xls = pd.ExcelFile(DATA_FILE)
    bat_dfs, bowl_dfs, field_dfs = [], [], []
    raw_names = set()
    
    for s in xls.sheet_names:
        header_idx = 1 if 'practice' in s.lower() else 0
        df = pd.read_excel(xls, sheet_name=s, header=header_idx)
        df.columns = [clean_col_name(c) for c in df.columns]
        if 'Player' in df.columns:
            df['Player'] = df['Player'].apply(clean_prefix)
            raw_names.update(df['Player'].dropna().unique())
            
            if 'bat' in s.lower(): bat_dfs.append(df)
            elif 'bowl' in s.lower(): bowl_dfs.append(df)
            
            f_cols = [c for c in df.columns if any(x in c.lower() for x in ['catch', 'run out', 'stump', 'fielding'])]
            if f_cols:
                f_df = df[['Player'] + f_cols].copy()
                field_dfs.append(f_df)

    raw_bat = pd.concat(bat_dfs, ignore_index=True) if bat_dfs else pd.DataFrame()
    raw_bowl = pd.concat(bowl_dfs, ignore_index=True) if bowl_dfs else pd.DataFrame()
    raw_field = pd.concat(field_dfs, ignore_index=True) if field_dfs else pd.DataFrame()
    
    return raw_bat, raw_bowl, raw_field, sorted(list(raw_names))

@st.cache_data(show_spinner="Crunching AI Ratings with Custom Weights...")
def calculate_ratings(raw_bat, raw_bowl, raw_field, mapping, w):
    if raw_bat.empty and raw_bowl.empty: return pd.DataFrame()
    
    rbat, rbowl, rfield = raw_bat.copy(), raw_bowl.copy(), raw_field.copy()

    if not rbat.empty: rbat['Player'] = rbat['Player'].map(mapping).fillna(rbat['Player'])
    if not rbowl.empty: rbowl['Player'] = rbowl['Player'].map(mapping).fillna(rbowl['Player'])
    if not rfield.empty: rfield['Player'] = rfield['Player'].map(mapping).fillna(rfield['Player'])

    fours_col = next((c for c in rbat.columns if str(c).lower().strip() in ['4s', 'fours', '4', "4's"]), None)
    sixes_col = next((c for c in rbat.columns if str(c).lower().strip() in ['6s', 'sixes', '6', "6's"]), None)

    for col in ['Runs', 'SR', 'Inns', 'NO']:
        rbat[col] = pd.to_numeric(rbat.get(col, 0), errors='coerce').fillna(0)
    
    if fours_col: rbat[fours_col] = pd.to_numeric(rbat[fours_col], errors='coerce').fillna(0)
    if sixes_col: rbat[sixes_col] = pd.to_numeric(rbat[sixes_col], errors='coerce').fillna(0)
        
    rbat['Balls_Faced'] = np.where(rbat['SR'] > 0, (rbat['Runs'] / rbat['SR']) * 100, 0)
    rbat['Dismissals'] = (rbat['Inns'] - rbat['NO']).clip(lower=0)
    
    bat_agg_kwargs = {
        'Inns': ('Inns', 'sum'),
        'Runs_bat': ('Runs', 'sum'),
        'Balls_Faced': ('Balls_Faced', 'sum'),
        'Dismissals': ('Dismissals', 'sum')
    }
    if fours_col: bat_agg_kwargs['Fours'] = (fours_col, 'sum')
    if sixes_col: bat_agg_kwargs['Sixes'] = (sixes_col, 'sum')
    
    agg_bat = rbat.groupby('Player').agg(**bat_agg_kwargs).reset_index()

    extras = np.zeros(len(rbowl))
    if not rbowl.empty:
        for col in rbowl.columns:
            cl = str(col).lower().strip()
            if cl in ['wd', 'wide', 'wides', 'nb', 'no ball', 'noballs', 'no balls']:
                extras += pd.to_numeric(rbowl[col], errors='coerce').fillna(0).values
        rbowl['Total_Extras'] = extras

        for col in ['Runs', 'Wkts', 'Overs', 'Total_Extras']:
            rbowl[col] = pd.to_numeric(rbowl.get(col, 0), errors='coerce').fillna(0)
        rbowl['Balls_Bowled'] = rbowl['Overs'].apply(overs_to_balls)
        
        agg_bowl = rbowl.groupby('Player').agg(Balls_Bowled=('Balls_Bowled', 'sum'), Runs_bowl=('Runs', 'sum'), Wkts=('Wkts', 'sum'), Total_Extras=('Total_Extras', 'sum')).reset_index()
    else:
        agg_bowl = pd.DataFrame(columns=['Player', 'Balls_Bowled', 'Runs_bowl', 'Wkts', 'Total_Extras'])

    if not rfield.empty:
        for col in rfield.columns:
            if col != 'Player': rfield[col] = pd.to_numeric(rfield[col], errors='coerce').fillna(0)
        rfield['Total_Fielding'] = rfield.drop(columns=['Player']).sum(axis=1)
        agg_field = rfield.groupby('Player').agg(Total_Fielding=('Total_Fielding', 'sum')).reset_index()
    else:
        agg_field = pd.DataFrame(columns=['Player', 'Total_Fielding'])

    valid = pd.merge(agg_bat, agg_bowl, on='Player', how='outer')
    valid = pd.merge(valid, agg_field, on='Player', how='outer').fillna(0)
    
    valid = valid[(valid['Balls_Faced'] >= 12) | (valid['Balls_Bowled'] >= 18)].copy()
    if valid.empty: return pd.DataFrame()

    if 'Fours' not in valid.columns: valid['Fours'] = 0
    if 'Sixes' not in valid.columns: valid['Sixes'] = 0
    
    valid['Bound_Runs'] = (valid['Fours'] * 4) + (valid['Sixes'] * 6)
    valid['Boundary_Pct'] = np.where(valid['Runs_bat'] > 0, (valid['Bound_Runs'] / valid['Runs_bat']) * 100, 0)
    valid['SR_bat'] = np.where(valid['Balls_Faced'] > 0, (valid['Runs_bat'] / valid['Balls_Faced']) * 100, 0)
    valid['Econ'] = np.where(valid['Balls_Bowled'] > 0, (valid['Runs_bowl'] / valid['Balls_Bowled']) * 6, 999)
    valid['Bat Avg'] = np.where(valid['Dismissals'] > 0, valid['Runs_bat'] / valid['Dismissals'], valid['Runs_bat'])
    valid['Bowl Avg'] = np.where(valid['Wkts'] > 0, valid['Runs_bowl'] / valid['Wkts'], 0)
    valid['Bowl SR'] = np.where(valid['Wkts'] > 0, valid['Balls_Bowled'] / valid['Wkts'], valid['Balls_Bowled'])
    valid['Bat BPD'] = np.where(valid['Dismissals'] > 0, valid['Balls_Faced'] / valid['Dismissals'], valid['Balls_Faced'])
    valid['Extras_Rate'] = np.where(valid['Balls_Bowled'] > 0, (valid['Total_Extras'] / valid['Balls_Bowled']) * 6, 0)
    valid['Overs'] = valid['Balls_Bowled'].apply(format_overs)

    mean_avg = valid['Runs_bat'].sum() / (valid['Dismissals'].sum() or 1)
    mean_bpd = valid['Balls_Faced'].sum() / (valid['Dismissals'].sum() or 1)
    mean_bowl_sr = valid['Balls_Bowled'].sum() / (valid['Wkts'].sum() or 1)
    mean_extras_rate = (valid['Total_Extras'].sum() / (valid['Balls_Bowled'].sum() or 1)) * 6

    valid['Sm_Avg'] = (valid['Runs_bat'] + (mean_avg * 3)) / (valid['Dismissals'] + 3)
    valid['Sm_SR'] = ((valid['Runs_bat'] + ((valid['Runs_bat'].sum() / (valid['Balls_Faced'].sum() or 1) * 100) / 100 * 30)) / (valid['Balls_Faced'] + 30)) * 100
    valid['Sm_BPD'] = (valid['Balls_Faced'] + (mean_bpd * 3)) / (valid['Dismissals'] + 3)
    
    league_bound_pct = valid['Bound_Runs'].sum() / (valid['Runs_bat'].sum() or 1)
    valid['Sm_Boundary'] = ((valid['Bound_Runs'] + (league_bound_pct * 50)) / (valid['Runs_bat'] + 50)) * 100
    
    valid['Sm_Econ'] = ((valid['Runs_bowl'] + ((valid['Runs_bowl'].sum() / (valid['Balls_Bowled'].sum() or 1) * 6) / 6 * 30)) / (valid['Balls_Bowled'] + 30)) * 6
    valid['Sm_Avg_bowl'] = (valid['Runs_bowl'] + ((valid['Runs_bowl'].sum() / (valid['Balls_Bowled'].sum() or 1) * 6) * 5)) / (valid['Wkts'] + 5)
    valid['Sm_Bowl_SR'] = (valid['Balls_Bowled'] + (mean_bowl_sr * 5)) / (valid['Wkts'] + 5)
    valid['Sm_Extras'] = ((valid['Total_Extras'] + (mean_extras_rate / 6 * 30)) / (valid['Balls_Bowled'] + 30)) * 6

    z_runs = (valid['Runs_bat'] - valid['Runs_bat'].mean()) / (valid['Runs_bat'].std() or 1)
    z_avg = (valid['Sm_Avg'] - valid['Sm_Avg'].mean()) / (valid['Sm_Avg'].std() or 1)
    z_sr = (valid['Sm_SR'] - valid['Sm_SR'].mean()) / (valid['Sm_SR'].std() or 1)
    z_bpd = (valid['Sm_BPD'] - valid['Sm_BPD'].mean()) / (valid['Sm_BPD'].std() or 1)
    z_bound = (valid['Sm_Boundary'] - valid['Sm_Boundary'].mean()) / (valid['Sm_Boundary'].std() or 1)
    
    z_wkts = (valid['Wkts'] - valid['Wkts'].mean()) / (valid['Wkts'].std() or 1)
    z_econ = (valid['Sm_Econ'].mean() - valid['Sm_Econ']) / (valid['Sm_Econ'].std() or 1)
    z_avg_bowl = (valid['Sm_Avg_bowl'].mean() - valid['Sm_Avg_bowl']) / (valid['Sm_Avg_bowl'].std() or 1)
    z_bowl_sr = (valid['Sm_Bowl_SR'].mean() - valid['Sm_Bowl_SR']) / (valid['Sm_Bowl_SR'].std() or 1) 
    z_extras = (valid['Sm_Extras'].mean() - valid['Sm_Extras']) / (valid['Sm_Extras'].std() or 1) 
    
    valid['Fielding_Score'] = (valid['Total_Fielding'] - valid['Total_Fielding'].mean()) / (valid['Total_Fielding'].std() or 1)

    bat_tot = w['bat_runs'] + w['bat_avg'] + w['bat_sr'] + w['bat_bpd'] + w['bat_bound'] or 1
    bowl_tot = w['bowl_wkts'] + w['bowl_econ'] + w['bowl_avg'] + w['bowl_sr'] + w['bowl_extras'] or 1
    
    valid['Bat_Score'] = (z_runs * (w['bat_runs']/bat_tot)) + (z_avg * (w['bat_avg']/bat_tot)) + (z_sr * (w['bat_sr']/bat_tot)) + (z_bpd * (w['bat_bpd']/bat_tot)) + (z_bound * (w['bat_bound']/bat_tot))
    valid['Bowl_Score'] = (z_wkts * (w['bowl_wkts']/bowl_tot)) + (z_econ * (w['bowl_econ']/bowl_tot)) + (z_avg_bowl * (w['bowl_avg']/bowl_tot)) + (z_bowl_sr * (w['bowl_sr']/bowl_tot)) + (z_extras * (w['bowl_extras']/bowl_tot))
    valid['Boundary_Score'] = z_bound 

    valid['Role'] = valid.apply(lambda r: 'All-Rounder' if r['Balls_Faced'] >= 15 and r['Balls_Bowled'] >= 18 else ('Bowler' if r['Balls_Bowled'] >= 18 else 'Batter'), axis=1)
    
    def calc_final(r):
        if r['Role'] == 'Batter': 
            tot = w['wt_batter_bat'] + w['wt_batter_field'] or 1
            return (r['Bat_Score'] * (w['wt_batter_bat']/tot)) + (r['Fielding_Score'] * (w['wt_batter_field']/tot))
        elif r['Role'] == 'Bowler': 
            tot = w['wt_bowler_bowl'] + w['wt_bowler_field'] or 1
            return (r['Bowl_Score'] * (w['wt_bowler_bowl']/tot)) + (r['Fielding_Score'] * (w['wt_bowler_field']/tot))
        else: 
            tot = w['wt_ar_bat'] + w['wt_ar_bowl'] + w['wt_ar_field'] or 1
            base = (r['Bat_Score'] * (w['wt_ar_bat']/tot)) + (r['Bowl_Score'] * (w['wt_ar_bowl']/tot)) + (r['Fielding_Score'] * (w['wt_ar_field']/tot))
            return base * w['ar_multiplier']
        
    valid['Final_Raw'] = valid.apply(calc_final, axis=1)

    mean_raw = valid['Final_Raw'].mean()
    std_raw = valid['Final_Raw'].std() or 1
    
    valid['AI Rating'] = 20.0 + ((valid['Final_Raw'] - mean_raw) / std_raw) * 3.33
    valid['AI Rating'] = valid['AI Rating'].clip(lower=10.0, upper=30.0).round(1)
    
    for col, new_col in [('Bat_Score', 'Bat_Rating'), ('Bowl_Score', 'Bowl_Rating'), ('Fielding_Score', 'Field_Rating')]:
        mean_val = valid[col].mean()
        std_val = valid[col].std() or 1
        valid[new_col] = 20.0 + ((valid[col] - mean_val) / std_val) * 3.33
        valid[new_col] = valid[new_col].clip(lower=10.0, upper=30.0).round(1)

    valid['Tier'] = pd.cut(valid['AI Rating'], bins=[0, 16.9, 22.9, 26.9, 31], labels=["Bronze", "Silver", "Gold", "Platinum"])
    return valid

file_time = os.path.getmtime(DATA_FILE) if os.path.exists(DATA_FILE) else 0
raw_bat_cache, raw_bowl_cache, raw_field_cache, all_raw_names = get_raw_excel_data(file_time)
master_df = calculate_ratings(raw_bat_cache, raw_bowl_cache, raw_field_cache, saved_mapping, algo_weights)

if not master_df.empty:
    def get_avg_scout(player_name):
        evals = human_ratings.get(player_name, {})
        if not evals: return None
        total, count = 0, 0
        for e, data in evals.items():
            if isinstance(data, dict) and "Final" in data:
                total += data["Final"]
                count += 1
        return round(total / count, 1) if count > 0 else None
    
    master_df['Avg Scout Score'] = master_df['Player'].apply(get_avg_scout)
    master_df['Draft Status'] = master_df['Player'].apply(lambda x: draft_state.get(x, "Available"))
    master_df['Leadership'] = master_df['Player'].apply(lambda x: leadership_state.get(x, "None"))
    master_df['Pool Status'] = master_df['Player'].apply(lambda x: pool_state.get(x, "Team Player"))

# --- UI TABS ---
tab1, tab2, tab3, tab4, tab5, tab6, tab7 = st.tabs(["🏆 Live Draft Board", "🎯 Snake Draft Room", "📊 Team Analytics", "🕸️ Player Profiles", "📝 Committee Scouting", "🧠 Methodology", "⚙️ Admin & Data"])

# --- TAB 1: DRAFT BOARD ---
with tab1:
    if master_df.empty:
        st.info("👋 Welcome! Please navigate to the '⚙️ Admin & Data' tab and upload your Excel stats file.")
    else:
        st.subheader("Live Interactive Draft Board")
        
        c1, c2, c3, c4 = st.columns(4)
        c1.metric("Total Players", len(master_df))
        c2.metric("Available Players", len(master_df[master_df['Draft Status'] == "Available"]))
        c3.metric("Platinum Tier", len(master_df[master_df['Tier'] == "Platinum"]))
        c4.metric("Avg AI Rating", round(master_df['AI Rating'].mean(), 1))

        f1, f2 = st.columns(2)
        role_f = f1.selectbox("Filter Role", ["All", "Batter", "Bowler", "All-Rounder"])
        status_f = f2.selectbox("Filter Status", ["Available Only", "All Players"] + [t for t in TEAMS if t != "Available"])
        
        disp_df = master_df.copy()
        if role_f != "All": disp_df = disp_df[disp_df['Role'] == role_f]
        if status_f == "Available Only": disp_df = disp_df[disp_df['Draft Status'] == "Available"]
        elif status_f != "All Players": disp_df = disp_df[disp_df['Draft Status'] == status_f]

        disp_df['Player'] = disp_df.apply(lambda r: f"{r['Player']} (C)" if r['Leadership'] == 'Captain' else (f"{r['Player']} (VC)" if r['Leadership'] == 'Vice Captain' else r['Player']), axis=1)

        col_order = [
            'Player', 'Role', 'Tier', 'Pool Status',
            'Runs_bat', 'Bat Avg', 'SR_bat', 'Boundary_Pct',
            'Wkts', 'Bowl Avg', 'Bowl SR', 'Econ', 'Extras_Rate',
            'Total_Fielding', 
            'Bat_Rating', 'Bowl_Rating', 'Field_Rating', 'Avg Scout Score', 'AI Rating', 
            'Draft Status'
        ]
        disp_df = disp_df[col_order].sort_values('AI Rating', ascending=False)
        
        disp_df.columns = [
            'Player', 'Role', 'Tier', 'Pool Status',
            'Runs', 'Bat Avg', 'Bat SR', 'Bound %',
            'Wkts', 'Bowl Avg', 'Bowl SR', 'Econ', 'Extras/Ov',
            'Fielding', 
            'Bat Rtg', 'Bowl Rtg', 'Field Rtg', 'Scout Rating', 'AI Rating', 
            'Draft Status'
        ]
        
        disp_df = disp_df.set_index('Player')
        
        styled_df = disp_df.style.background_gradient(subset=['AI Rating'], cmap='RdYlGn', vmin=10, vmax=30)\
            .format({
                'Runs': '{:.0f}', 'Wkts': '{:.0f}', 'Fielding': '{:.0f}',
                'Bat Avg': '{:.2f}', 'Bat SR': '{:.1f}', 'Bound %': '{:.1f}%',
                'Bowl Avg': '{:.2f}', 'Bowl SR': '{:.1f}', 'Econ': '{:.2f}', 'Extras/Ov': '{:.2f}',
                'Bat Rtg': '{:.1f}', 'Bowl Rtg': '{:.1f}', 'Field Rtg': '{:.1f}',
                'Scout Rating': '{:.1f}', 'AI Rating': '{:.1f}'
            }, na_rep="-")
            
        st.dataframe(styled_df, use_container_width=True)

# --- TAB 2: SNAKE DRAFT ROOM ---
with tab2:
    if master_df.empty:
        st.info("Data required.")
    else:
        st.subheader("🎯 Live Snake Draft & Salary Cap Room")
        
        squad_size = int(algo_weights.get("squad_size", 11))
        team_budget = float(algo_weights.get("team_budget", 240.0))
        
        valid_teams = [t for t in TEAMS if t != "Available"]
        num_teams = len(valid_teams)
        drafted_df = master_df[master_df['Draft Status'] != "Available"]
        num_drafted = len(drafted_df)
        
        round_num = (num_drafted // num_teams) + 1
        pick_in_round = num_drafted % num_teams
        if (round_num % 2) != 0:
            on_the_clock = valid_teams[pick_in_round]
        else:
            on_the_clock = valid_teams[num_teams - 1 - pick_in_round]
            
        st.markdown("---")
        st.markdown(f"### 🟢 ON THE CLOCK: **{on_the_clock}** (Round {round_num}, Pick {pick_in_round + 1})")
        
        # --- STYLED SUMMARY STATISTICS TABLE ---
        st.markdown("#### 💰 Franchise Cap & Roster Summary")
        sum_df = pd.DataFrame(index=["Total Points Burnt", "Total Players Added", "Total Points Allocated", "Remaining Points", "Players yet to take"])
        
        for t in valid_teams:
            t_df = drafted_df[drafted_df['Draft Status'] == t]
            spent = t_df['AI Rating'].sum() if not t_df.empty else 0.0
            rem = team_budget - spent
            players_added = len(t_df)
            players_needed = squad_size - players_added
            
            sum_df[t] = ["", "", "", "", ""]
            sum_df[f"{t} Rtg"] = [
                f"{spent:.1f}", str(players_added), f"{team_budget:.1f}", f"{rem:.1f}", str(players_needed)
            ]
            
        def color_summary(df):
            style_df = pd.DataFrame('', index=df.index, columns=df.columns)
            for c in df.columns:
                if "Rtg" in c:
                    try:
                        rem_val = float(df.at["Remaining Points", c])
                        style_df.at["Remaining Points", c] = 'background-color: #d4edda; color: #155724; font-weight: bold;' if rem_val >= 0 else 'background-color: #f8d7da; color: #721c24; font-weight: bold;'
                        style_df.at["Total Points Burnt", c] = 'background-color: #fff3cd; color: #856404; font-weight: bold;'
                        style_df.at["Total Points Allocated", c] = 'background-color: #e2e3e5; color: #383d41;'
                    except:
                        pass
            return style_df

        st.dataframe(sum_df.style.apply(color_summary, axis=None), use_container_width=True)

        # --- EXCEL-STYLE INTERACTIVE DRAFT GRID ---
        st.markdown("#### 📋 Official Draft Board Grid")
        st.write("Click an empty cell to pick a player. If you need to remove someone or made a mistake, select **`--- CLEAR PICK ---`**.")
        
        # Display Errors safely and clear immediately
        if "draft_error" in st.session_state:
            st.error(st.session_state.draft_error)
            del st.session_state.draft_error

        grid_df = pd.DataFrame(index=[f"Round {i+1}" for i in range(squad_size)])
        for t in valid_teams:
            grid_df[t] = ""
            grid_df[f"{t} Rtg"] = np.nan

        # Pre-fill grid with 🔒 + Name + Rating 
        for t in valid_teams:
            t_players = [p for p, team in draft_state.items() if team == t]
            for i, p_name in enumerate(t_players):
                if i < squad_size:
                    p_match = master_df[master_df['Player'] == p_name]
                    if not p_match.empty:
                        p_rtg = float(p_match.iloc[0]['AI Rating'])
                        # The exact format injected into the backend grid DataFrame
                        grid_df.iat[i, grid_df.columns.get_loc(t)] = f"🔒 {p_name} ({p_rtg:.1f})"
                        grid_df.iat[i, grid_df.columns.get_loc(f"{t} Rtg")] = p_rtg

        # Build Clean Dropdown Options
        avail_team_players = master_df[(master_df['Draft Status'] == "Available") & (master_df['Pool Status'] == "Team Player")].sort_values('AI Rating', ascending=False)
        avail_opts = [f"{row['Player']} ({row['AI Rating']:.1f})" for _, row in avail_team_players.iterrows()]

        col_config = {}
        for t in valid_teams:
            t_drafted = [p for p, team in draft_state.items() if team == t]
            t_drafted_opts = []
            
            for p in t_drafted:
                p_match = master_df[master_df['Player'] == p]
                if not p_match.empty:
                    # These locked options match the grid values EXACTLY so Streamlit renders them properly
                    t_drafted_opts.append(f"🔒 {p} ({p_match.iloc[0]['AI Rating']:.1f})")
            
            # Add CLEAR PICK tool to perfectly clear a cell without fighting Streamlit's backspace logic
            opts = ["--- CLEAR PICK ---"] + t_drafted_opts + avail_opts
            col_config[t] = st.column_config.SelectboxColumn(f"{t} Name", options=opts, required=False)
            col_config[f"{t} Rtg"] = st.column_config.Column("Ratings", disabled=True)

        # Apply Heatmap STRICTLY to Rating columns
        rtg_cols = [f"{t} Rtg" for t in valid_teams]
        styled_grid = grid_df.style.background_gradient(subset=rtg_cols, cmap='RdYlGn', vmin=10, vmax=30).format({c: "{:.1f}" for c in rtg_cols}, na_rep="")

        edited_grid = st.data_editor(styled_grid, column_config=col_config, use_container_width=True, key="live_grid")

        # Catching edits from the Grid
        grid_changed = False
        replacements = {}
        deletions = set()
        additions = []

        for t in valid_teams:
            for i in range(squad_size):
                old_val = grid_df.iat[i, grid_df.columns.get_loc(t)]
                new_val = edited_grid.iat[i, edited_grid.columns.get_loc(t)]
                
                old_v = extract_name(old_val)
                new_v = extract_name(new_val)
                
                if old_v != new_v:
                    grid_changed = True
                    if old_v and new_v:
                        replacements[old_v] = (new_v, t)
                    elif old_v and not new_v:
                        deletions.add(old_v)
                    elif not old_v and new_v:
                        additions.append((new_v, t))

        if grid_changed:
            # 1. Validation Loop: Check constraints before saving
            for t in valid_teams:
                t_players = []
                for i in range(squad_size):
                    val = edited_grid.iat[i, edited_grid.columns.get_loc(t)]
                    clean_name = extract_name(val)
                    if clean_name: 
                        t_players.append(clean_name)
                
                # Validation: Duplicate checks
                if len(t_players) != len(set(t_players)):
                    for item in t_players:
                        if t_players.count(item) > 1:
                            st.session_state.draft_error = f"❌ REJECTED: You attempted to pick '{item}' multiple times. Edit reverted."
                            break
                    if "live_grid" in st.session_state: del st.session_state["live_grid"]
                    st.rerun()

                # Validation: Cap Math
                t_spent = sum([master_df[master_df['Player'] == p]['AI Rating'].iloc[0] for p in t_players if p in master_df['Player'].values])
                t_rem = team_budget - t_spent
                t_count = len(t_players)
                
                if t_spent > team_budget:
                    st.session_state.draft_error = f"❌ SALARY CAP EXCEEDED: {t} cannot afford this roster! Edit reverted."
                    if "live_grid" in st.session_state: del st.session_state["live_grid"]
                    st.rerun()

                if t_rem < (10.0 * (squad_size - t_count)):
                    min_req = 10.0 * (squad_size - t_count)
                    st.session_state.draft_error = f"❌ INVALID ROSTER: {t} must save at least {min_req:.1f} points for remaining {squad_size - t_count} slots. Edit reverted."
                    if "live_grid" in st.session_state: del st.session_state["live_grid"]
                    st.rerun()
            
            # 2. Rebuild draft_state carefully to PRESERVE ROUND ORDER
            new_draft_state = {}
            for p, team in draft_state.items():
                if p in deletions:
                    continue
                if p in replacements:
                    new_p, new_t = replacements[p]
                    new_draft_state[new_p] = new_t
                else:
                    new_draft_state[p] = team
            
            for p, t in additions:
                new_draft_state[p] = t
                            
            save_json(DRAFT_FILE, new_draft_state)
            
            # FORCE clear the frontend cache so it rebuilds smoothly without hanging
            if "live_grid" in st.session_state: 
                del st.session_state["live_grid"]
            st.rerun()

# --- TAB 3: TEAM ANALYTICS ---
with tab3:
    if master_df.empty:
        st.info("Data required.")
    else:
        st.subheader("Live Team Balance Analytics")
        drafted = master_df[master_df['Draft Status'] != "Available"]
        if drafted.empty:
            st.info("No players drafted yet. Use the Draft Room to begin.")
        else:
            team_stats = drafted.groupby('Draft Status').agg(
                Players=('Player', 'count'), Total_AI_Rating=('AI Rating', 'sum'),
                Avg_AI_Rating=('AI Rating', 'mean'), Total_Runs=('Runs_bat', 'sum'), Total_Wkts=('Wkts', 'sum')
            ).reset_index()
            
            c1, c2 = st.columns(2)
            fig1 = px.bar(team_stats, x='Draft Status', y='Avg_AI_Rating', title="Average Team Rating", color='Draft Status')
            c1.plotly_chart(fig1, use_container_width=True)
            
            fig2 = px.bar(team_stats, x='Draft Status', y='Total_AI_Rating', title="Total Team Power Score", color='Draft Status')
            c2.plotly_chart(fig2, use_container_width=True)
            
            st.dataframe(team_stats.style.format({'Avg_AI_Rating': "{:.1f}"}), use_container_width=True)

# --- TAB 4: PLAYER PROFILES ---
with tab4:
    if master_df.empty:
        st.info("Data required.")
    else:
        st.subheader("Advanced Player Scouting Profiles")
        selected_player = st.selectbox("Search Player", master_df.sort_values('AI Rating', ascending=False)['Player'])
        
        if selected_player:
            p_data = master_df[master_df['Player'] == selected_player].iloc[0]
            tag = "🏆 *(Captain)*" if p_data['Leadership'] == 'Captain' else ("⭐ *(Vice Captain)*" if p_data['Leadership'] == 'Vice Captain' else "")
            
            pc1, pc2 = st.columns([1, 2])
            with pc1:
                st.markdown(f"### {p_data['Player']} {tag}")
                st.markdown(f"**Role:** {p_data['Role']} | **Tier:** {p_data['Tier']} | **Status:** {p_data['Pool Status']}")
                st.markdown(f"**AI Rating:** {p_data['AI Rating']:.1f}/30.0")
                st.markdown(f"**Drafted To:** {p_data['Draft Status']}")
                st.markdown("---")
                st.markdown(f"**Total Runs:** {int(p_data['Runs_bat'])} *(Avg: {p_data['Bat Avg']:.2f}, SR: {p_data['SR_bat']:.1f}, Bound %: {p_data['Boundary_Pct']:.1f}%)*")
                st.markdown(f"**Total Wkts:** {int(p_data['Wkts'])} *(Avg: {p_data['Bowl Avg']:.2f}, SR: {p_data['Bowl SR']:.1f}, Econ: {p_data['Econ']:.2f})*")
                st.markdown(f"**Fielding Dismissals:** {int(p_data['Total_Fielding'])}")
                
                st.markdown("---")
                st.markdown("**Committee Score Breakdown:**")
                scores = human_ratings.get(selected_player, {})
                if not scores: 
                    st.write("*No committee reviews yet.*")
                else:
                    for evaluator, data in scores.items():
                        if isinstance(data, dict):
                            st.write(f"- **{evaluator} ({data.get('Final', 0):.1f})** | *{data.get('Role')} | Bat: {data.get('Bat')}, Bowl: {data.get('Bowl')}, Field: {data.get('Field')}*")
                        else:
                            st.write(f"- **{evaluator}**: Legacy Score Ignored")

            with pc2:
                categories = ['Batting Volume', 'Strike Rate', 'Boundary Threat', 'Wicket Taking', 'Economy (Reversed)', 'Bowling Discipline', 'Fielding Impact']
                r_bat = (p_data['Bat_Score'] - master_df['Bat_Score'].min()) / (master_df['Bat_Score'].max() - master_df['Bat_Score'].min() + 0.01)
                r_sr = (p_data['SR_bat'] - master_df['SR_bat'].min()) / (master_df['SR_bat'].max() - master_df['SR_bat'].min() + 0.01)
                r_bound = (p_data['Boundary_Score'] - master_df['Boundary_Score'].min()) / (master_df['Boundary_Score'].max() - master_df['Boundary_Score'].min() + 0.01)
                
                r_bowl = (p_data['Bowl_Score'] - master_df['Bowl_Score'].min()) / (master_df['Bowl_Score'].max() - master_df['Bowl_Score'].min() + 0.01)
                r_econ = 1 - ((p_data['Econ'] - master_df['Econ'].min()) / (master_df['Econ'].max() - master_df['Econ'].min() + 0.01))
                if p_data['Econ'] == 0 or p_data['Econ'] == 999: r_econ = 0
                
                r_disc = 1 - ((p_data['Extras_Rate'] - master_df['Extras_Rate'].min()) / (master_df['Extras_Rate'].max() - master_df['Extras_Rate'].min() + 0.01))
                if p_data['Extras_Rate'] == 0 and master_df['Extras_Rate'].max() == 0: r_disc = 1
                
                r_field = (p_data['Fielding_Score'] - master_df['Fielding_Score'].min()) / (master_df['Fielding_Score'].max() - master_df['Fielding_Score'].min() + 0.01)
                
                fig = go.Figure()
                fig.add_trace(go.Scatterpolar(
                    r=[r_bat, r_sr, r_bound, r_bowl, r_econ, r_disc, r_field, r_bat], theta=categories + [categories[0]], fill='toself', line_color='orange'
                ))
                fig.update_layout(polar=dict(radialaxis=dict(visible=False, range=[0, 1])), showlegend=False, title="Skill Heptagon")
                st.plotly_chart(fig, use_container_width=True)

# --- TAB 5: COMMITTEE SCOUTING ---
with tab5:
    if master_df.empty:
        st.info("Data required.")
    else:
        st.subheader("📝 Committee Scouting (10-30 Scale)")
        st.write("Rate the player's disciplines. The engine will calculate the final score automatically.")
        
        sc1, sc2 = st.columns(2)
        evaluator = sc1.selectbox("Select Your Name", auth_users)
        scout_player = sc2.selectbox("Select Player to Rate", master_df.sort_values('Player')['Player'])
        
        existing_data = human_ratings.get(scout_player, {}).get(evaluator, {})
        if not isinstance(existing_data, dict): existing_data = {}
        
        engine_role = "Batter"
        if scout_player in master_df['Player'].values:
            engine_role = master_df[master_df['Player'] == scout_player].iloc[0]['Role']
            
        def_role = existing_data.get("Role", engine_role)
        def_bat = existing_data.get("Bat", 20.0)
        def_bowl = existing_data.get("Bowl", 20.0)
        def_field = existing_data.get("Field", 20.0)
        
        sel_role = st.selectbox("Assign Player Role", ["Batter", "Bowler", "All-Rounder"], index=["Batter", "Bowler", "All-Rounder"].index(def_role))
        
        st.markdown("##### Assign Skill Ratings & Compare with AI Peers")
        c_bat, c_bowl, c_fld = st.columns(3)
        
        def display_peer_slider(col_obj, title, def_val, skill_col):
            with col_obj:
                val = st.slider(title, 10.0, 30.0, float(def_val), step=0.5)
                min_v, max_v = val - 1.5, val + 1.5
                peers = master_df[(master_df[skill_col] >= min_v) & (master_df[skill_col] <= max_v)].copy()
                
                if peers.empty:
                    st.caption(f"*No AI peers in {min_v:.1f} - {max_v:.1f} range.*")
                else:
                    peers['diff'] = abs(peers[skill_col] - val)
                    top_peers = peers.sort_values('diff').head(6)
                    peer_str = ", ".join([f"{r['Player']} ({r[skill_col]:.1f})" for _, r in top_peers.iterrows()])
                    st.caption(f"**🔍 AI Peers ({min_v:.1f}-{max_v:.1f}):**<br>{peer_str}", unsafe_allow_html=True)
                return val

        val_bat = display_peer_slider(c_bat, "Batting Rating", def_bat, "Bat_Rating")
        val_bowl = display_peer_slider(c_bowl, "Bowling Rating", def_bowl, "Bowl_Rating")
        val_fld = display_peer_slider(c_fld, "Fielding Rating", def_field, "Field_Rating")
        
        w = algo_weights
        if sel_role == 'Batter': 
            tot = w['wt_batter_bat'] + w['wt_batter_field'] or 1
            calc_final = (val_bat * (w['wt_batter_bat']/tot)) + (val_fld * (w['wt_batter_field']/tot))
        elif sel_role == 'Bowler': 
            tot = w['wt_bowler_bowl'] + w['wt_bowler_field'] or 1
            calc_final = (val_bowl * (w['wt_bowler_bowl']/tot)) + (val_fld * (w['wt_bowler_field']/tot))
        else: 
            tot = w['wt_ar_bat'] + w['wt_ar_bowl'] + w['wt_ar_field'] or 1
            raw_ar = (val_bat * (w['wt_ar_bat']/tot)) + (val_bowl * (w['wt_ar_bowl']/tot)) + (val_fld * (w['wt_ar_field']/tot))
            calc_final = min(30.0, raw_ar * w['ar_multiplier'])
            
        st.info(f"**Calculated Final Scout Rating:** {calc_final:.1f} / 30.0")
        
        if st.button("Save Rating", type="primary"):
            if scout_player not in human_ratings: human_ratings[scout_player] = {}
            human_ratings[scout_player][evaluator] = {
                "Role": sel_role,
                "Bat": val_bat,
                "Bowl": val_bowl,
                "Field": val_fld,
                "Final": float(round(calc_final, 1))
            }
            save_json(RATINGS_FILE, human_ratings)
            st.success(f"Score detailed breakdown saved for {scout_player}!")
            st.rerun()

# --- TAB 6: METHODOLOGY ---
with tab6:
    w = algo_weights
    st.subheader("🧠 How the AI Rating is Calculated")
    
    st.markdown(f"""
    To create maximum separation between players and ensure a fair draft, the BPL Engine uses dynamic, customizable weighting logic.

    ### 1. Granular Core Metrics
    The engine balances **8 distinct data points** dynamically based on your custom Admin settings:
    *   **Batting Score:** Total Runs ({w['bat_runs']}%), Batting Avg ({w['bat_avg']}%), Strike Rate ({w['bat_sr']}%), Balls Per Dismissal ({w['bat_bpd']}%), and Boundary Impact ({w['bat_bound']}%).
    *   **Bowling Score:** Total Wickets ({w['bowl_wkts']}%), Economy Rate ({w['bowl_econ']}%), Bowling Avg ({w['bowl_avg']}%), Bowling Strike Rate ({w['bowl_sr']}%), and Extras/Discipline Penalty ({w['bowl_extras']}%).
    *   **Fielding Score:** The sum of all Catches, Run-Outs, and Stumpings.

    ### 2. Weights & The All-Rounder Premium
    Players are assigned a role based on strict minimum thresholds:
    * **Batter:** Batting ({w['wt_batter_bat']}%) + Fielding ({w['wt_batter_field']}%)
    * **Bowler:** Bowling ({w['wt_bowler_bowl']}%) + Fielding ({w['wt_bowler_field']}%)
    * **All-Rounder:** `[Batting ({w['wt_ar_bat']}%) + Bowling ({w['wt_ar_bowl']}%) + Fielding ({w['wt_ar_field']}%)] * {w['ar_multiplier']} Multiplier`
    
    ### 3. T-Score Distribution
    The absolute league average player is hardcoded to receive exactly a **20.0 AI Rating**. The algorithm applies a 3.33 standard deviation spread, naturally fanning the players out across the 10.0 to 30.0 range.
    """)

# --- TAB 7: ADMIN & DATA ---
with tab7:
    st.subheader("⚙️ System Management")
    
    password_attempt = st.text_input("Enter Admin Password to unlock controls", type="password")
    
    if password_attempt == ADMIN_PASSWORD:
        st.success("Admin Access Granted")
        st.markdown("---")
        
        ac1, ac2 = st.columns(2)
        with ac1:
            st.markdown("**1. Setup Rosters, Leadership & Player Pool**")
            st.caption("Assign Captains and toggle whether a player is eligible for the draft (Team Player) or held in reserve (Pool Player).")
            if not master_df.empty:
                draft_df = pd.DataFrame({
                    "Player": master_df['Player'], 
                    "Draft Status": master_df['Draft Status'],
                    "Leadership": master_df['Leadership'],
                    "Pool Status": master_df['Pool Status']
                })
                edited_draft = st.data_editor(
                    draft_df, 
                    column_config={
                        "Draft Status": st.column_config.SelectboxColumn(options=TEAMS),
                        "Leadership": st.column_config.SelectboxColumn(options=["None", "Captain", "Vice Captain"]),
                        "Pool Status": st.column_config.SelectboxColumn(options=["Team Player", "Pool Player"])
                    }, 
                    hide_index=True, use_container_width=True
                )
                if st.button("💾 Save Rosters, Roles & Pool"):
                    new_draft = dict(zip(edited_draft["Player"], edited_draft["Draft Status"]))
                    new_pool = dict(zip(edited_draft["Player"], edited_draft["Pool Status"]))
                    
                    new_leaders = dict(zip(edited_draft["Player"], edited_draft["Leadership"]))
                    new_leaders = {k: v for k, v in new_leaders.items() if v != "None"}
                    
                    save_json(DRAFT_FILE, new_draft)
                    save_json(LEADERSHIP_FILE, new_leaders)
                    save_json(POOL_FILE, new_pool)
                    st.success("Draft, Leadership, and Pool updated!")
                    st.rerun()
            else:
                st.info("Upload data first.")

        with ac2:
            st.markdown("**2. Authorized Evaluators**")
            users_text = st.text_area("List names (comma separated)", ", ".join(auth_users))
            if st.button("Update Evaluators"):
                new_users = [u.strip() for u in users_text.split(",")]
                save_json(USERS_FILE, new_users)
                st.success("Evaluators updated!")
                st.rerun()
            
            st.markdown("---")
            st.markdown("**3. Data File Upload**")
            uploaded_file = st.file_uploader("Upload Raw Stats (.xlsx)", type=["xlsx"])
            if uploaded_file:
                with open(DATA_FILE, "wb") as f: f.write(uploaded_file.getbuffer())
                st.success("Uploaded!")
                st.rerun()

        st.markdown("---")
        st.markdown("### 4. 🎛️ Draft & Algorithm Settings")
        st.write("Modify the mathematical importance of each metric, or configure the Salary Cap.")
        
        st.markdown("##### 🎯 Salary Cap Rules")
        st.write("*(Changes made here save instantly and reflect in the Draft Room!)*")
        sc1, sc2, sc3 = st.columns(3)
        w_squad_size = sc1.number_input("Max Players Per Team", value=int(algo_weights.get("squad_size", 11)), step=1, key="cap_squad", on_change=update_cap_settings)
        w_team_budget = sc2.number_input("Team Point Budget", value=float(algo_weights.get("team_budget", 240.0)), step=5.0, key="cap_budget", on_change=update_cap_settings)
        
        if not master_df.empty:
            valid_teams_ct = len([t for t in TEAMS if t != "Available"])
            ideal_pool = master_df[(master_df['Pool Status'] == "Team Player")].sort_values('AI Rating', ascending=False).head(valid_teams_ct * w_squad_size)
            suggested_cap = ideal_pool['AI Rating'].sum() / valid_teams_ct
            sc3.info(f"**Suggested Cap:** {suggested_cap:.1f} pts\n\n*(Based on top {valid_teams_ct * w_squad_size} Team Players)*")
        
        st.markdown("##### 🏏 Batting Metrics (%)")
        b1, b2, b3, b4, b5 = st.columns(5)
        w_bat_runs = b1.number_input("Total Runs", value=float(algo_weights["bat_runs"]))
        w_bat_avg = b2.number_input("Batting Avg", value=float(algo_weights["bat_avg"]))
        w_bat_sr = b3.number_input("Strike Rate", value=float(algo_weights["bat_sr"]))
        w_bat_bpd = b4.number_input("Balls/Dismissal", value=float(algo_weights["bat_bpd"]))
        w_bat_bound = b5.number_input("Boundary %", value=float(algo_weights["bat_bound"]))
        
        st.markdown("##### ⚾ Bowling Metrics (%)")
        bo1, bo2, bo3, bo4, bo5 = st.columns(5)
        w_bowl_wkts = bo1.number_input("Total Wickets", value=float(algo_weights["bowl_wkts"]))
        w_bowl_econ = bo2.number_input("Economy", value=float(algo_weights["bowl_econ"]))
        w_bowl_avg = bo3.number_input("Bowling Avg", value=float(algo_weights["bowl_avg"]))
        w_bowl_sr = bo4.number_input("Bowling SR", value=float(algo_weights["bowl_sr"]))
        w_bowl_extras = bo5.number_input("Extras Penalty", value=float(algo_weights["bowl_extras"]))
        
        st.markdown("##### ⚖️ Role Distribution (%) & Multiplier")
        r1, r2, r3 = st.columns(3)
        with r1:
            st.markdown("**Pure Batter**")
            w_wt_batter_bat = st.number_input("Batting % (Batter)", value=float(algo_weights["wt_batter_bat"]))
            w_wt_batter_field = st.number_input("Fielding % (Batter)", value=float(algo_weights["wt_batter_field"]))
        with r2:
            st.markdown("**Pure Bowler**")
            w_wt_bowler_bowl = st.number_input("Bowling % (Bowler)", value=float(algo_weights["wt_bowler_bowl"]))
            w_wt_bowler_field = st.number_input("Fielding % (Bowler)", value=float(algo_weights["wt_bowler_field"]))
        with r3:
            st.markdown("**All-Rounder**")
            w_wt_ar_bat = st.number_input("Batting % (AR)", value=float(algo_weights["wt_ar_bat"]))
            w_wt_ar_bowl = st.number_input("Bowling % (AR)", value=float(algo_weights["wt_ar_bowl"]))
            w_wt_ar_field = st.number_input("Fielding % (AR)", value=float(algo_weights["wt_ar_field"]))
            w_ar_multiplier = st.number_input("AR Multiplier", value=float(algo_weights["ar_multiplier"]), step=0.1)

        if st.button("⚙️ Save Custom Settings & Recalculate"):
            new_weights = {
                "squad_size": w_squad_size, "team_budget": w_team_budget,
                "bat_runs": w_bat_runs, "bat_avg": w_bat_avg, "bat_sr": w_bat_sr, "bat_bpd": w_bat_bpd, "bat_bound": w_bat_bound,
                "bowl_wkts": w_bowl_wkts, "bowl_econ": w_bowl_econ, "bowl_avg": w_bowl_avg, "bowl_sr": w_bowl_sr, "bowl_extras": w_bowl_extras,
                "wt_batter_bat": w_wt_batter_bat, "wt_batter_field": w_wt_batter_field,
                "wt_bowler_bowl": w_wt_bowler_bowl, "wt_bowler_field": w_wt_bowler_field,
                "wt_ar_bat": w_wt_ar_bat, "wt_ar_bowl": w_wt_ar_bowl, "wt_ar_field": w_wt_ar_field, "ar_multiplier": w_ar_multiplier
            }
            save_json(WEIGHTS_FILE, new_weights)
            st.success("✅ Engine settings updated! Head to the Draft Room to see the new Cap limits.")
            st.rerun()

        st.markdown("---")
        st.markdown("### 5. 🛠️ Player Name Aliases & Merge Tool")
        st.write("Edit the **'Merged Name'** column to merge aliases dynamically.")
        if all_raw_names:
            mapping_records = [{"Original Name": n, "Merged Name": saved_mapping.get(n, n)} for n in all_raw_names]
            mapping_df = pd.DataFrame(mapping_records)
            edited_mapping = st.data_editor(mapping_df, use_container_width=True, hide_index=True)
            
            if st.button("💾 Save Player Mappings Permanently"):
                new_map = dict(zip(edited_mapping["Original Name"], edited_mapping["Merged Name"]))
                saved_mapping.update(new_map)
                save_json(MAPPING_FILE, saved_mapping)
                st.success("✅ Mappings saved permanently! The engine will now recalculate.")
                st.rerun()
        else:
            st.info("Upload an Excel file to start mapping names.")

        st.markdown("---")
        st.markdown("### 6. 💾 Permanent Cloud Backup & Restore")
        st.write("Because free servers reset when code changes, download your server state to save your mappings and rosters permanently.")
        
        bc1, bc2 = st.columns(2)
        with bc1:
            backup_data = {
                "mappings": load_json(MAPPING_FILE, DEFAULT_MAPPINGS),
                "draft": load_json(DRAFT_FILE, {}),
                "ratings": load_json(RATINGS_FILE, {}),
                "weights": load_json(WEIGHTS_FILE, DEFAULT_WEIGHTS),
                "leadership": load_json(LEADERSHIP_FILE, {}),
                "pool": load_json(POOL_FILE, {})
            }
            backup_json = json.dumps(backup_data, indent=2).encode('utf-8')
            st.download_button(
                label="📥 Download Server Backup", 
                data=backup_json, 
                file_name="bpl_server_backup.json", 
                mime="application/json",
                type="primary"
            )
            
        with bc2:
            restore_file = st.file_uploader("📤 Restore from Backup (.json)", type=["json"])
            if restore_file:
                restore_data = json.load(restore_file)
                if "mappings" in restore_data: save_json(MAPPING_FILE, restore_data.get("mappings", {}))
                if "draft" in restore_data: save_json(DRAFT_FILE, restore_data.get("draft", {}))
                if "ratings" in restore_data: save_json(RATINGS_FILE, restore_data.get("ratings", {}))
                if "weights" in restore_data: save_json(WEIGHTS_FILE, restore_data.get("weights", {}))
                if "leadership" in restore_data: save_json(LEADERSHIP_FILE, restore_data.get("leadership", {}))
                if "pool" in restore_data: save_json(POOL_FILE, restore_data.get("pool", {}))
                st.success("✅ Server state fully restored!")
                st.rerun()
