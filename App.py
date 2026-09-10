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

ADMIN_PASSWORD = "bpladmin" 

DEFAULT_MAPPINGS = {
    "Praveenraj Starkey": "Praveenraj Starkey (Merged)",
    "Praveen Starkey": "Praveenraj Starkey (Merged)",
    "Godwin .": "Godwin (Merged)",
    "Godwin M": "Godwin (Merged)"
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

# --- LOAD STATES ---
saved_mapping = load_json(MAPPING_FILE, {})
mapping_changed = False
for original, merged in DEFAULT_MAPPINGS.items():
    if original not in saved_mapping:
        saved_mapping[original] = merged
        mapping_changed = True
if mapping_changed:
    save_json(MAPPING_FILE, saved_mapping)

draft_state = load_json(DRAFT_FILE, {}) 
human_ratings = load_json(RATINGS_FILE, {}) 
auth_users = load_json(USERS_FILE, ["Admin", "Captain 1", "Captain 2"])
TEAMS = ["Available", "Team 1", "Team 2", "Team 3", "Team 4", "Team 5"]

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
            
            # Look for fielding data across any sheet
            f_cols = [c for c in df.columns if any(x in c.lower() for x in ['catch', 'run out', 'stump', 'fielding'])]
            if f_cols:
                f_df = df[['Player'] + f_cols].copy()
                field_dfs.append(f_df)

    raw_bat = pd.concat(bat_dfs, ignore_index=True) if bat_dfs else pd.DataFrame()
    raw_bowl = pd.concat(bowl_dfs, ignore_index=True) if bowl_dfs else pd.DataFrame()
    raw_field = pd.concat(field_dfs, ignore_index=True) if field_dfs else pd.DataFrame()
    
    return raw_bat, raw_bowl, raw_field, sorted(list(raw_names))

@st.cache_data(show_spinner="Crunching AI Ratings...")
def calculate_ratings(raw_bat, raw_bowl, raw_field, mapping):
    if raw_bat.empty and raw_bowl.empty: return pd.DataFrame()
    
    rbat, rbowl, rfield = raw_bat.copy(), raw_bowl.copy(), raw_field.copy()

    if not rbat.empty: rbat['Player'] = rbat['Player'].map(mapping).fillna(rbat['Player'])
    if not rbowl.empty: rbowl['Player'] = rbowl['Player'].map(mapping).fillna(rbowl['Player'])
    if not rfield.empty: rfield['Player'] = rfield['Player'].map(mapping).fillna(rfield['Player'])

    # Aggregate Batting
    for col in ['Runs', 'SR', 'Inns', 'NO']:
        rbat[col] = pd.to_numeric(rbat.get(col, 0), errors='coerce').fillna(0)
    rbat['Balls_Faced'] = np.where(rbat['SR'] > 0, (rbat['Runs'] / rbat['SR']) * 100, 0)
    rbat['Dismissals'] = (rbat['Inns'] - rbat['NO']).clip(lower=0)
    agg_bat = rbat.groupby('Player').agg(Inns=('Inns', 'sum'), Runs_bat=('Runs', 'sum'), Balls_Faced=('Balls_Faced', 'sum'), Dismissals=('Dismissals', 'sum')).reset_index()

    # Aggregate Bowling
    for col in ['Runs', 'Wkts', 'Overs']:
        rbowl[col] = pd.to_numeric(rbowl.get(col, 0), errors='coerce').fillna(0)
    rbowl['Balls_Bowled'] = rbowl['Overs'].apply(overs_to_balls)
    agg_bowl = rbowl.groupby('Player').agg(Balls_Bowled=('Balls_Bowled', 'sum'), Runs_bowl=('Runs', 'sum'), Wkts=('Wkts', 'sum')).reset_index()

    # Aggregate Fielding
    if not rfield.empty:
        for col in rfield.columns:
            if col != 'Player': rfield[col] = pd.to_numeric(rfield[col], errors='coerce').fillna(0)
        rfield['Total_Fielding'] = rfield.drop(columns=['Player']).sum(axis=1)
        agg_field = rfield.groupby('Player').agg(Total_Fielding=('Total_Fielding', 'sum')).reset_index()
    else:
        agg_field = pd.DataFrame(columns=['Player', 'Total_Fielding'])

    # Merge Disciplines
    valid = pd.merge(agg_bat, agg_bowl, on='Player', how='outer')
    valid = pd.merge(valid, agg_field, on='Player', how='outer').fillna(0)
    
    # Filter valid players
    valid = valid[(valid['Balls_Faced'] >= 12) | (valid['Balls_Bowled'] >= 18)].copy()
    if valid.empty: return pd.DataFrame()

    # Standardize Stats
    valid['SR_bat'] = np.where(valid['Balls_Faced'] > 0, (valid['Runs_bat'] / valid['Balls_Faced']) * 100, 0)
    valid['Econ'] = np.where(valid['Balls_Bowled'] > 0, (valid['Runs_bowl'] / valid['Balls_Bowled']) * 6, 999)
    valid['Bat Avg'] = np.where(valid['Dismissals'] > 0, valid['Runs_bat'] / valid['Dismissals'], valid['Runs_bat'])
    valid['Bowl Avg'] = np.where(valid['Wkts'] > 0, valid['Runs_bowl'] / valid['Wkts'], 0)
    valid['Overs'] = valid['Balls_Bowled'].apply(format_overs)

    # Bayesian Smoothing
    mean_avg = valid['Runs_bat'].sum() / (valid['Dismissals'].sum() or 1)
    valid['Sm_Avg'] = (valid['Runs_bat'] + (mean_avg * 3)) / (valid['Dismissals'] + 3)
    valid['Sm_SR'] = ((valid['Runs_bat'] + ((valid['Runs_bat'].sum() / (valid['Balls_Faced'].sum() or 1) * 100) / 100 * 30)) / (valid['Balls_Faced'] + 30)) * 100
    valid['Sm_Econ'] = ((valid['Runs_bowl'] + ((valid['Runs_bowl'].sum() / (valid['Balls_Bowled'].sum() or 1) * 6) / 6 * 30)) / (valid['Balls_Bowled'] + 30)) * 6
    valid['Sm_Avg_bowl'] = (valid['Runs_bowl'] + ((valid['Runs_bowl'].sum() / (valid['Balls_Bowled'].sum() or 1) * 6) * 5)) / (valid['Wkts'] + 5)

    # Z-Scores
    z_runs = (valid['Runs_bat'] - valid['Runs_bat'].mean()) / (valid['Runs_bat'].std() or 1)
    z_avg = (valid['Sm_Avg'] - valid['Sm_Avg'].mean()) / (valid['Sm_Avg'].std() or 1)
    z_sr = (valid['Sm_SR'] - valid['Sm_SR'].mean()) / (valid['Sm_SR'].std() or 1)
    valid['Bat_Score'] = z_runs * 0.3 + z_avg * 0.4 + z_sr * 0.3

    z_wkts = (valid['Wkts'] - valid['Wkts'].mean()) / (valid['Wkts'].std() or 1)
    z_econ = (valid['Sm_Econ'].mean() - valid['Sm_Econ']) / (valid['Sm_Econ'].std() or 1)
    z_avg_bowl = (valid['Sm_Avg_bowl'].mean() - valid['Sm_Avg_bowl']) / (valid['Sm_Avg_bowl'].std() or 1)
    valid['Bowl_Score'] = z_wkts * 0.4 + z_econ * 0.35 + z_avg_bowl * 0.25
    
    valid['Fielding_Score'] = (valid['Total_Fielding'] - valid['Total_Fielding'].mean()) / (valid['Total_Fielding'].std() or 1)

    # Role & Final Logic (with Fielding Weights)
    valid['Role'] = valid.apply(lambda r: 'All-Rounder' if r['Balls_Faced'] >= 15 and r['Balls_Bowled'] >= 18 else ('Bowler' if r['Balls_Bowled'] >= 18 else 'Batter'), axis=1)
    
    def calc_final(r):
        if r['Role'] == 'Batter': return (r['Bat_Score'] * 0.85) + (r['Fielding_Score'] * 0.15)
        elif r['Role'] == 'Bowler': return (r['Bowl_Score'] * 0.85) + (r['Fielding_Score'] * 0.15)
        else: return ((r['Bat_Score'] * 0.425) + (r['Bowl_Score'] * 0.425) + (r['Fielding_Score'] * 0.15)) * 1.3
        
    valid['Final_Raw'] = valid.apply(calc_final, axis=1)

    min_raw, max_raw = np.percentile(valid['Final_Raw'], 1), np.percentile(valid['Final_Raw'], 99)
    if max_raw == min_raw: max_raw = min_raw + 1 
    
    valid['AI Rating'] = ((valid['Final_Raw'] - min_raw) / (max_raw - min_raw)) * 20.0 + 10.0
    valid['AI Rating'] = valid['AI Rating'].clip(lower=10.0, upper=30.0).round(1)
    valid['Tier'] = pd.cut(valid['AI Rating'], bins=[0, 18, 23, 27, 31], labels=["Bronze", "Silver", "Gold", "Platinum"])
    
    return valid

# Execution Engine
file_time = os.path.getmtime(DATA_FILE) if os.path.exists(DATA_FILE) else 0
raw_bat_cache, raw_bowl_cache, raw_field_cache, all_raw_names = get_raw_excel_data(file_time)
master_df = calculate_ratings(raw_bat_cache, raw_bowl_cache, raw_field_cache, saved_mapping)

if not master_df.empty:
    def get_avg_scout(player_name):
        scores = human_ratings.get(player_name, {}).values()
        return round(sum(scores)/len(scores), 1) if scores else None
    
    master_df['Avg Scout Score'] = master_df['Player'].apply(get_avg_scout)
    master_df['Draft Status'] = master_df['Player'].apply(lambda x: draft_state.get(x, "Available"))

# --- UI TABS ---
tab1, tab2, tab3, tab4, tab5, tab6 = st.tabs(["🏆 Live Draft Board", "📊 Team Analytics", "🕸️ Player Profiles", "📝 Committee Scouting", "🧠 Methodology", "⚙️ Admin & Data"])

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

        # Formatting Output Columns
        disp_df = disp_df[['Player', 'Role', 'Tier', 'Runs_bat', 'Bat Avg', 'Wkts', 'Bowl Avg', 'Econ', 'Total_Fielding', 'AI Rating', 'Avg Scout Score', 'Draft Status']].sort_values('AI Rating', ascending=False)
        disp_df.columns = ['Player', 'Role', 'Tier', 'Runs', 'Bat Avg', 'Wkts', 'Bowl Avg', 'Econ', 'Fielding', 'AI Rating', 'Scout Rating', 'Draft Status']
        
        styled_df = disp_df.style.background_gradient(subset=['AI Rating'], cmap='RdYlGn', vmin=10, vmax=30)\
            .format({
                'Runs': '{:.0f}',
                'Wkts': '{:.0f}',
                'Fielding': '{:.0f}',
                'Bat Avg': '{:.2f}',
                'Bowl Avg': '{:.2f}',
                'Econ': '{:.2f}',
                'AI Rating': '{:.1f}',
                'Scout Rating': '{:.1f}'
            }, na_rep="-")
            
        st.dataframe(styled_df, use_container_width=True, hide_index=True)

# --- TAB 2: TEAM ANALYTICS ---
with tab2:
    if master_df.empty:
        st.info("Data required.")
    else:
        st.subheader("Live Team Balance Analytics")
        drafted = master_df[master_df['Draft Status'] != "Available"]
        if drafted.empty:
            st.info("No players drafted yet. Update Draft Status in the Admin tab.")
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

# --- TAB 3: PLAYER PROFILES ---
with tab3:
    if master_df.empty:
        st.info("Data required.")
    else:
        st.subheader("Advanced Player Scouting Profiles")
        selected_player = st.selectbox("Search Player", master_df.sort_values('AI Rating', ascending=False)['Player'])
        
        if selected_player:
            p_data = master_df[master_df['Player'] == selected_player].iloc[0]
            
            pc1, pc2 = st.columns([1, 2])
            with pc1:
                st.markdown(f"### {p_data['Player']}")
                st.markdown(f"**Role:** {p_data['Role']} | **Tier:** {p_data['Tier']}")
                st.markdown(f"**AI Rating:** {p_data['AI Rating']:.1f}/30.0")
                st.markdown(f"**Drafted To:** {p_data['Draft Status']}")
                st.markdown("---")
                st.markdown(f"**Total Runs:** {int(p_data['Runs_bat'])} *(Avg: {p_data['Bat Avg']:.2f}, SR: {p_data['SR_bat']:.1f})*")
                st.markdown(f"**Total Wkts:** {int(p_data['Wkts'])} *(Avg: {p_data['Bowl Avg']:.2f}, Econ: {p_data['Econ']:.2f}, Overs: {p_data['Overs']})*")
                st.markdown(f"**Fielding Dismissals:** {int(p_data['Total_Fielding'])}")
                
                st.markdown("---")
                st.markdown("**Committee Scores:**")
                scores = human_ratings.get(selected_player, {})
                if not scores: st.write("*No committee reviews yet.*")
                for evaluator, score in scores.items():
                    st.write(f"- {evaluator}: {score:.1f}/10")

            with pc2:
                categories = ['Batting Volume', 'Strike Rate', 'Wicket Taking', 'Economy (Reversed)', 'Fielding Impact']
                r_bat = (p_data['Bat_Score'] - master_df['Bat_Score'].min()) / (master_df['Bat_Score'].max() - master_df['Bat_Score'].min() + 0.01)
                r_sr = (p_data['SR_bat'] - master_df['SR_bat'].min()) / (master_df['SR_bat'].max() - master_df['SR_bat'].min() + 0.01)
                r_bowl = (p_data['Bowl_Score'] - master_df['Bowl_Score'].min()) / (master_df['Bowl_Score'].max() - master_df['Bowl_Score'].min() + 0.01)
                r_econ = 1 - ((p_data['Econ'] - master_df['Econ'].min()) / (master_df['Econ'].max() - master_df['Econ'].min() + 0.01))
                if p_data['Econ'] == 0 or p_data['Econ'] == 999: r_econ = 0
                r_field = (p_data['Fielding_Score'] - master_df['Fielding_Score'].min()) / (master_df['Fielding_Score'].max() - master_df['Fielding_Score'].min() + 0.01)
                
                fig = go.Figure()
                fig.add_trace(go.Scatterpolar(
                    r=[r_bat, r_sr, r_bowl, r_econ, r_field, r_bat], theta=categories + [categories[0]], fill='toself', line_color='orange'
                ))
                fig.update_layout(polar=dict(radialaxis=dict(visible=False, range=[0, 1])), showlegend=False, title="Skill Polygon")
                st.plotly_chart(fig, use_container_width=True)

# --- TAB 4: COMMITTEE SCOUTING ---
with tab4:
    if master_df.empty:
        st.info("Data required.")
    else:
        st.subheader("📝 Individual Committee Ratings")
        st.write("Authorized users can assign their personal scouting scores (1-10) to players.")
        
        sc1, sc2 = st.columns(2)
        evaluator = sc1.selectbox("Select Your Name", auth_users)
        scout_player = sc2.selectbox("Select Player to Rate", master_df.sort_values('Player')['Player'])
        
        current_score = human_ratings.get(scout_player, {}).get(evaluator, 5.0)
        new_score = st.slider("Assign Rating (1 = Poor, 10 = Elite)", 1.0, 10.0, float(current_score), step=0.5)
        
        if st.button("Save Rating", type="primary"):
            if scout_player not in human_ratings: human_ratings[scout_player] = {}
            human_ratings[scout_player][evaluator] = new_score
            save_json(RATINGS_FILE, human_ratings)
            st.success(f"Score saved for {scout_player}!")
            st.rerun()

# --- TAB 5: METHODOLOGY ---
with tab5:
    st.subheader("🧠 How the AI Rating is Calculated")
    
    st.markdown("""
    To ensure fair valuations for the draft, the BPL Engine uses a robust **Z-Score Normalization** model mapped to a **10.0 – 30.0 scale**. The engine automatically parses batting, bowling, and fielding metrics across all uploaded sheets.

    ### 1. The Core Metrics
    * **Batting Score:** Calculated using Total Runs (30%), Batting Average (40%), and Strike Rate (30%).
    * **Bowling Score:** Calculated using Total Wickets (40%), Economy Rate (35%), and Bowling Average (25%).
    * **Fielding Score:** The engine sums all Catches, Run-Outs, and Stumpings a player is involved in across all scorecards.
    
    ### 2. Bayesian Smoothing (The "Reality Check")
    If a player scores 12 runs off 2 balls and never gets out, their Strike Rate is technically 600.0 and their Average is Infinity. If a bowler bowls 1 over and takes 1 wicket for 2 runs, their Economy is 2.00. 
    
    To prevent these tiny sample sizes from breaking the ratings, the engine injects "fictitious" baseline stats (the league average) into every single player's batting and bowling records. This gently pulls outliers back down to reality, while rewarding players who maintain excellent stats over a *large volume* of matches.
    
    ### 3. Z-Scores (Comparing to the League)
    Instead of using raw numbers, the engine converts every smoothed stat into a **Z-Score**. A Z-Score measures exactly how many standard deviations a player is above or below the league average (+1.5 means you are significantly better than the average player, -0.5 means you are slightly below average).
    
    ### 4. Weights & Role Designation
    Players are assigned a role based on strict minimum thresholds. The Final Raw Score is then calculated using precise weights:
    * **Batter (Faced 15+ balls):** Batting Score (85%) + Fielding Score (15%)
    * **Bowler (Bowled 18+ balls):** Bowling Score (85%) + Fielding Score (15%)
    * **All-Rounder (Met BOTH thresholds):** `[(Batting * 42.5%) + (Bowling * 42.5%) + (Fielding * 15%)] * 1.3 Multiplier`
    
    *The 1.3x All-Rounder Multiplier reflects the immense tactical value of a player who saves a roster spot by contributing elite skill in both innings.*

    ---
    ### 📖 Worked Example: Calculating "Player X"
    Let's look at how the math actually applies to a hypothetical All-Rounder in the BPL.

    Assume the **League Averages** are currently:
    *   **Batting:** Average = 15.00, Strike Rate = 110.0
    *   **Bowling:** Economy = 8.50, Average = 20.00
    *   **Fielding:** Total Dismissals = 1.0 (with a standard deviation of 1.0)

    **Player X's Raw Stats:**
    *   **Batting:** 120 Runs, 4 Dismissals, 80 Balls Faced (Raw Avg: 30.00, Raw SR: 150.0)
    *   **Bowling:** 8 Wickets, 16 Overs (96 balls), 112 Runs Given (Raw Econ: 7.00, Raw Avg: 14.00)
    *   **Fielding:** 3 Catches

    **Step 1: Smoothing**
    The engine adds the league average to Player X's stats.
    *   *Smoothed Batting Avg* drops slightly from 30.00 down to **~23.50**.
    *   *Smoothed Economy* rises slightly from 7.00 up to **~7.35**.

    **Step 2: Z-Score Math**
    The engine calculates how Player X compares to the rest of the league:
    *   *Z_Runs:* +1.50, *Z_Bat_Avg:* +1.20, *Z_SR:* +1.80
    *   **Batting Score** = `(1.50 * 0.3) + (1.20 * 0.4) + (1.80 * 0.3)` = **+1.47**

    *   *Z_Wkts:* +1.40, *Z_Econ:* +1.10 (Reversed), *Z_Bowl_Avg:* +1.30 (Reversed)
    *   **Bowling Score** = `(1.40 * 0.4) + (1.10 * 0.35) + (1.30 * 0.25)` = **+1.27**

    *   *Z_Fielding:* +2.00 (Since 3 catches is 2 standard deviations above the league average of 1)
    *   **Fielding Score** = **+2.00**

    **Step 3: Weighting & The All-Rounder Multiplier**
    Because Player X faced >15 balls AND bowled >18 balls, they qualify for the 42.5/42.5/15 weight split and the premium multiplier.
    *   Combined Base = `(1.47 * 0.425) + (1.27 * 0.425) + (2.00 * 0.15)`
    *   Combined Base = `0.62 + 0.54 + 0.30` = **1.46**
    *   Final Raw Score = `1.46 * 1.3` = **1.90**

    **Step 4: Mapping to the 10-30 Scale**
    The engine analyzes the highest and lowest scores across the entire BPL. A Final Raw Score of 1.90 represents massive overall impact. The formula maps this to the visual scale, giving Player X an **AI Rating of 26.8**, placing them firmly in the **Platinum Tier**.
    """)

# --- TAB 6: ADMIN & DATA ---
with tab6:
    st.subheader("⚙️ System Management")
    
    password_attempt = st.text_input("Enter Admin Password to unlock controls", type="password")
    
    if password_attempt == ADMIN_PASSWORD:
        st.success("Admin Access Granted")
        st.markdown("---")
        
        ac1, ac2 = st.columns(2)
        with ac1:
            st.markdown("**1. Live Draft Management**")
            if not master_df.empty:
                draft_df = pd.DataFrame({"Player": master_df['Player'], "Draft Status": master_df['Draft Status']})
                edited_draft = st.data_editor(draft_df, column_config={"Draft Status": st.column_config.SelectboxColumn(options=TEAMS)}, hide_index=True, use_container_width=True)
                if st.button("💾 Save Draft Rosters"):
                    new_draft = dict(zip(edited_draft["Player"], edited_draft["Draft Status"]))
                    save_json(DRAFT_FILE, new_draft)
                    st.success("Draft updated!")
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
        st.markdown("### 4. 🛠️ Player Name Aliases & Merge Tool")
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
        st.markdown("### 5. 💾 Permanent Cloud Backup & Restore")
        st.write("Because free servers reset when code changes, download your server state to save your mappings and draft rosters permanently.")
        
        bc1, bc2 = st.columns(2)
        with bc1:
            backup_data = {
                "mappings": load_json(MAPPING_FILE, DEFAULT_MAPPINGS),
                "draft": load_json(DRAFT_FILE, {}),
                "ratings": load_json(RATINGS_FILE, {})
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
                if "mappings" in restore_data:
                    save_json(MAPPING_FILE, restore_data.get("mappings", {}))
                    save_json(DRAFT_FILE, restore_data.get("draft", {}))
                    save_json(RATINGS_FILE, restore_data.get("ratings", {}))
                    st.success("✅ Server state fully restored!")
                    st.rerun()
    elif password_attempt != "":
        st.error("Incorrect password.")
