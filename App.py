import streamlit as st
import pandas as pd
import numpy as np
import os
import json
import re
import plotly.graph_objects as go
import plotly.express as px

st.set_page_config(page_title="Tigers Cricket Club - Draft Engine", layout="wide", page_icon="🏏")

# --- FILE PATHS ---
DATA_FILE = "current_stats.xlsx"
MAPPING_FILE = "name_mapping.json"
DRAFT_FILE = "draft_state.json"
RATINGS_FILE = "human_ratings.json"
USERS_FILE = "authorized_users.json"

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
draft_state = load_json(DRAFT_FILE, {}) # {PlayerName: "Team 1"}
human_ratings = load_json(RATINGS_FILE, {}) # {PlayerName: {EvaluatorName: Rating}}
auth_users = load_json(USERS_FILE, ["Admin", "Captain 1", "Captain 2"])
TEAMS = ["Available", "Team 1", "Team 2", "Team 3", "Team 4", "Team 5"]

# --- HEADER ---
st.title("🏏 Tigers Cricket Club Draft Engine (Est. 2026)")
st.markdown("Advanced AI Rating, Live Roster Management, and Committee Scouting.")

# --- DATA PROCESSING ENGINE ---
@st.cache_data
def load_and_process_data(mapping):
    if not os.path.exists(DATA_FILE): return pd.DataFrame(), pd.DataFrame(), pd.DataFrame()
    xls = pd.ExcelFile(DATA_FILE)
    bat_dfs, bowl_dfs = [], []
    for s in xls.sheet_names:
        header_idx = 1 if 'practice' in s.lower() else 0
        df = pd.read_excel(xls, sheet_name=s, header=header_idx)
        df.columns = [clean_col_name(c) for c in df.columns]
        if 'Player' in df.columns:
            df['Player'] = df['Player'].apply(clean_prefix)
            if 'bat' in s.lower(): bat_dfs.append(df)
            elif 'bowl' in s.lower(): bowl_dfs.append(df)

    raw_bat = pd.concat(bat_dfs, ignore_index=True) if bat_dfs else pd.DataFrame()
    raw_bowl = pd.concat(bowl_dfs, ignore_index=True) if bowl_dfs else pd.DataFrame()
    
    if raw_bat.empty and raw_bowl.empty: return pd.DataFrame(), pd.DataFrame(), pd.DataFrame()

    if not raw_bat.empty: raw_bat['Player'] = raw_bat['Player'].map(mapping).fillna(raw_bat['Player'])
    if not raw_bowl.empty: raw_bowl['Player'] = raw_bowl['Player'].map(mapping).fillna(raw_bowl['Player'])

    for col in ['Runs', 'SR', 'Inns', 'NO']:
        raw_bat[col] = pd.to_numeric(raw_bat.get(col, 0), errors='coerce').fillna(0)
    raw_bat['Balls_Faced'] = np.where(raw_bat['SR'] > 0, (raw_bat['Runs'] / raw_bat['SR']) * 100, 0)
    raw_bat['Dismissals'] = (raw_bat['Inns'] - raw_bat['NO']).clip(lower=0)
    agg_bat = raw_bat.groupby('Player').agg(Inns=('Inns', 'sum'), Runs_bat=('Runs', 'sum'), Balls_Faced=('Balls_Faced', 'sum'), Dismissals=('Dismissals', 'sum')).reset_index()

    for col in ['Runs', 'Wkts', 'Overs']:
        raw_bowl[col] = pd.to_numeric(raw_bowl.get(col, 0), errors='coerce').fillna(0)
    raw_bowl['Balls_Bowled'] = raw_bowl['Overs'].apply(overs_to_balls)
    agg_bowl = raw_bowl.groupby('Player').agg(Balls_Bowled=('Balls_Bowled', 'sum'), Runs_bowl=('Runs', 'sum'), Wkts=('Wkts', 'sum')).reset_index()

    valid = pd.merge(agg_bat, agg_bowl, on='Player', how='outer').fillna(0)
    valid = valid[(valid['Balls_Faced'] >= 12) | (valid['Balls_Bowled'] >= 18)].copy()

    valid['SR_bat'] = np.where(valid['Balls_Faced'] > 0, (valid['Runs_bat'] / valid['Balls_Faced']) * 100, 0)
    valid['Econ'] = np.where(valid['Balls_Bowled'] > 0, (valid['Runs_bowl'] / valid['Balls_Bowled']) * 6, 999)
    valid['Bat Avg'] = np.where(valid['Dismissals'] > 0, valid['Runs_bat'] / valid['Dismissals'], valid['Runs_bat'])
    valid['Bowl Avg'] = np.where(valid['Wkts'] > 0, valid['Runs_bowl'] / valid['Wkts'], 0)
    valid['Overs'] = valid['Balls_Bowled'].apply(format_overs)

    mean_avg = valid['Runs_bat'].sum() / (valid['Dismissals'].sum() or 1)
    valid['Sm_Avg'] = (valid['Runs_bat'] + (mean_avg * 3)) / (valid['Dismissals'] + 3)
    valid['Sm_SR'] = ((valid['Runs_bat'] + ((valid['Runs_bat'].sum() / valid['Balls_Faced'].sum() * 100) / 100 * 30)) / (valid['Balls_Faced'] + 30)) * 100
    valid['Sm_Econ'] = ((valid['Runs_bowl'] + ((valid['Runs_bowl'].sum() / valid['Balls_Bowled'].sum() * 6) / 6 * 30)) / (valid['Balls_Bowled'] + 30)) * 6
    valid['Sm_Avg_bowl'] = (valid['Runs_bowl'] + ((valid['Runs_bowl'].sum() / valid['Balls_Bowled'].sum() * 6) * 5)) / (valid['Wkts'] + 5)

    z_runs = (valid['Runs_bat'] - valid['Runs_bat'].mean()) / valid['Runs_bat'].std()
    z_avg = (valid['Sm_Avg'] - valid['Sm_Avg'].mean()) / valid['Sm_Avg'].std()
    z_sr = (valid['Sm_SR'] - valid['Sm_SR'].mean()) / valid['Sm_SR'].std()
    valid['Bat_Score'] = z_runs * 0.3 + z_avg * 0.4 + z_sr * 0.3

    z_wkts = (valid['Wkts'] - valid['Wkts'].mean()) / valid['Wkts'].std()
    z_econ = (valid['Sm_Econ'].mean() - valid['Sm_Econ']) / valid['Sm_Econ'].std()
    z_avg_bowl = (valid['Sm_Avg_bowl'].mean() - valid['Sm_Avg_bowl']) / valid['Sm_Avg_bowl'].std()
    valid['Bowl_Score'] = z_wkts * 0.4 + z_econ * 0.35 + z_avg_bowl * 0.25

    valid['Role'] = valid.apply(lambda r: 'All-Rounder' if r['Balls_Faced'] >= 15 and r['Balls_Bowled'] >= 18 else ('Bowler' if r['Balls_Bowled'] >= 18 else 'Batter'), axis=1)
    valid['Final_Raw'] = valid.apply(lambda r: r['Bat_Score'] if r['Role'] == 'Batter' else (r['Bowl_Score'] if r['Role'] == 'Bowler' else (r['Bat_Score'] * 0.5 + r['Bowl_Score'] * 0.5) * 1.3), axis=1)

    min_raw, max_raw = np.percentile(valid['Final_Raw'], 1), np.percentile(valid['Final_Raw'], 99)
    valid['AI Rating'] = ((valid['Final_Raw'] - min_raw) / (max_raw - min_raw)) * 20.0 + 10.0
    valid['AI Rating'] = valid['AI Rating'].clip(lower=10.0, upper=30.0).round(1)

    # Tiers
    valid['Tier'] = pd.cut(valid['AI Rating'], bins=[0, 18, 23, 27, 31], labels=["Bronze", "Silver", "Gold", "Platinum"])
    
    return valid, raw_bat, raw_bowl

master_df, raw_bat, raw_bowl = load_and_process_data(saved_mapping)

if master_df.empty:
    st.info("👋 Welcome! Please navigate to the '⚙️ Admin' tab and upload your Excel stats file to initialize the engine.")
else:
    # Blend Human Ratings
    def get_avg_scout(player_name):
        scores = human_ratings.get(player_name, {}).values()
        return round(sum(scores)/len(scores), 1) if scores else None
    
    master_df['Avg Scout Score'] = master_df['Player'].apply(get_avg_scout)
    master_df['Draft Status'] = master_df['Player'].apply(lambda x: draft_state.get(x, "Available"))
    
    # --- UI TABS ---
    tab1, tab2, tab3, tab4, tab5 = st.tabs(["🏆 Live Draft Board", "📊 Team Analytics", "🕸️ Player Profiles", "📝 Committee Scouting", "⚙️ Admin & Data"])

    # --- TAB 1: DRAFT BOARD ---
    with tab1:
        st.subheader("Live Interactive Draft Board")
        
        # KPIs
        c1, c2, c3, c4 = st.columns(4)
        c1.metric("Total Players", len(master_df))
        c2.metric("Available Players", len(master_df[master_df['Draft Status'] == "Available"]))
        c3.metric("Platinum Tier", len(master_df[master_df['Tier'] == "Platinum"]))
        c4.metric("Avg AI Rating", round(master_df['AI Rating'].mean(), 1))

        # Filters
        f1, f2 = st.columns(2)
        role_f = f1.selectbox("Filter Role", ["All", "Batter", "Bowler", "All-Rounder"])
        status_f = f2.selectbox("Filter Status", ["Available Only", "All Players"] + [t for t in TEAMS if t != "Available"])
        
        disp_df = master_df.copy()
        if role_f != "All": disp_df = disp_df[disp_df['Role'] == role_f]
        if status_f == "Available Only": disp_df = disp_df[disp_df['Draft Status'] == "Available"]
        elif status_f != "All Players": disp_df = disp_df[disp_df['Draft Status'] == status_f]

        # Formatting for Display
        disp_df = disp_df[['Player', 'Role', 'Tier', 'Runs_bat', 'Wkts', 'Econ', 'AI Rating', 'Avg Scout Score', 'Draft Status']].sort_values('AI Rating', ascending=False)
        disp_df.columns = ['Player', 'Role', 'Tier', 'Runs', 'Wkts', 'Econ', 'AI Rating', 'Scout Rating', 'Draft Status']
        
        # Color coding Rating
        styled_df = disp_df.style.background_gradient(subset=['AI Rating'], cmap='RdYlGn', vmin=10, vmax=30)
        st.dataframe(styled_df, use_container_width=True, hide_index=True)

    # --- TAB 2: TEAM ANALYTICS ---
    with tab2:
        st.subheader("Live Team Balance Analytics")
        drafted = master_df[master_df['Draft Status'] != "Available"]
        if drafted.empty:
            st.info("No players drafted yet. Update Draft Status in the Admin tab.")
        else:
            team_stats = drafted.groupby('Draft Status').agg(
                Players=('Player', 'count'),
                Total_AI_Rating=('AI Rating', 'sum'),
                Avg_AI_Rating=('AI Rating', 'mean'),
                Total_Runs=('Runs_bat', 'sum'),
                Total_Wkts=('Wkts', 'sum')
            ).reset_index()
            
            c1, c2 = st.columns(2)
            fig1 = px.bar(team_stats, x='Draft Status', y='Avg_AI_Rating', title="Average Team Rating", color='Draft Status')
            c1.plotly_chart(fig1, use_container_width=True)
            
            fig2 = px.bar(team_stats, x='Draft Status', y='Total_AI_Rating', title="Total Team Power Score", color='Draft Status')
            c2.plotly_chart(fig2, use_container_width=True)
            
            st.dataframe(team_stats.style.format({'Avg_AI_Rating': "{:.1f}"}), use_container_width=True)

    # --- TAB 3: PLAYER PROFILES ---
    with tab3:
        st.subheader("Advanced Player Scouting Profiles")
        selected_player = st.selectbox("Search Player", master_df.sort_values('AI Rating', ascending=False)['Player'])
        
        if selected_player:
            p_data = master_df[master_df['Player'] == selected_player].iloc[0]
            
            pc1, pc2 = st.columns([1, 2])
            with pc1:
                st.markdown(f"### {p_data['Player']}")
                st.markdown(f"**Role:** {p_data['Role']} | **Tier:** {p_data['Tier']}")
                st.markdown(f"**AI Rating:** {p_data['AI Rating']}/30.0")
                st.markdown(f"**Drafted To:** {p_data['Draft Status']}")
                st.markdown("---")
                st.markdown(f"**Total Runs:** {int(p_data['Runs_bat'])} *(Avg: {p_data['Bat Avg']:.1f}, SR: {p_data['SR_bat']:.1f})*")
                st.markdown(f"**Total Wkts:** {int(p_data['Wkts'])} *(Econ: {p_data['Econ']:.1f}, Overs: {p_data['Overs']})*")
                
                # Show individual human ratings
                st.markdown("---")
                st.markdown("**Committee Scores:**")
                scores = human_ratings.get(selected_player, {})
                if not scores: st.write("*No committee reviews yet.*")
                for evaluator, score in scores.items():
                    st.write(f"- {evaluator}: {score}/10")

            with pc2:
                # Radar Chart Logic
                categories = ['Batting Volume', 'Strike Rate', 'Wicket Taking', 'Economy (Reversed)']
                
                # Normalize values purely for radar visualization (0 to 1 scale)
                r_bat = (p_data['Bat_Score'] - master_df['Bat_Score'].min()) / (master_df['Bat_Score'].max() - master_df['Bat_Score'].min() + 0.01)
                r_sr = (p_data['SR_bat'] - master_df['SR_bat'].min()) / (master_df['SR_bat'].max() - master_df['SR_bat'].min() + 0.01)
                r_bowl = (p_data['Bowl_Score'] - master_df['Bowl_Score'].min()) / (master_df['Bowl_Score'].max() - master_df['Bowl_Score'].min() + 0.01)
                # Reverse Econ for radar (lower econ = higher score on graph)
                r_econ = 1 - ((p_data['Econ'] - master_df['Econ'].min()) / (master_df['Econ'].max() - master_df['Econ'].min() + 0.01))
                if p_data['Econ'] == 0 or p_data['Econ'] == 999: r_econ = 0
                
                fig = go.Figure()
                fig.add_trace(go.Scatterpolar(
                    r=[r_bat, r_sr, r_bowl, r_econ, r_bat],
                    theta=categories + [categories[0]],
                    fill='toself', name=selected_player, line_color='orange'
                ))
                fig.update_layout(polar=dict(radialaxis=dict(visible=False, range=[0, 1])), showlegend=False, title="Skill Polygon")
                st.plotly_chart(fig, use_container_width=True)

    # --- TAB 4: COMMITTEE SCOUTING ---
    with tab4:
        st.subheader("📝 Individual Committee Ratings")
        st.write("Authorized users can assign their personal scouting scores (1-10) to players. This helps captains evaluate intangibles.")
        
        sc1, sc2 = st.columns(2)
        evaluator = sc1.selectbox("Select Your Name", auth_users)
        scout_player = sc2.selectbox("Select Player to Rate", master_df.sort_values('Player')['Player'])
        
        current_score = human_ratings.get(scout_player, {}).get(evaluator, 5)
        new_score = st.slider("Assign Rating (1 = Poor, 10 = Elite)", 1, 10, current_score)
        
        if st.button("Save Rating", type="primary"):
            if scout_player not in human_ratings: human_ratings[scout_player] = {}
            human_ratings[scout_player][evaluator] = new_score
            save_json(RATINGS_FILE, human_ratings)
            st.success(f"Score saved for {scout_player}!")
            st.rerun()

    # --- TAB 5: ADMIN & DATA ---
    with tab5:
        st.subheader("⚙️ System Management")
        
        ac1, ac2 = st.columns(2)
        with ac1:
            st.markdown("**1. Live Draft Management**")
            st.write("Assign players to teams to update the analytics tab.")
            draft_df = pd.DataFrame({"Player": master_df['Player'], "Draft Status": master_df['Draft Status']})
            edited_draft = st.data_editor(draft_df, column_config={"Draft Status": st.column_config.SelectboxColumn(options=TEAMS)}, hide_index=True, use_container_width=True)
            if st.button("💾 Save Draft Rosters"):
                new_draft = dict(zip(edited_draft["Player"], edited_draft["Draft Status"]))
                save_json(DRAFT_FILE, new_draft)
                st.success("Draft updated!")
                st.rerun()

        with ac2:
            st.markdown("**2. Authorized Committee Evaluators**")
            users_text = st.text_area("List names (comma separated)", ", ".join(auth_users))
            if st.button("Update Evaluators"):
                new_users = [u.strip() for u in users_text.split(",")]
                save_json(USERS_FILE, new_users)
                st.success("Evaluators updated!")
                st.rerun()
            
            st.markdown("---")
            st.markdown("**3. Data File Upload**")
            uploaded_file = st.file_uploader("Upload .xlsx", type=["xlsx"])
            if uploaded_file:
                with open(DATA_FILE, "wb") as f: f.write(uploaded_file.getbuffer())
                st.success("Uploaded!")
                st.rerun()
                
            if st.button("⚠️ Hard Reset Draft Status"):
                save_json(DRAFT_FILE, {})
                st.rerun()
