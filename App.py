import streamlit as st
import pandas as pd
import numpy as np

st.set_page_config(page_title="🏏 Cricket Draft Board & Rating Engine", layout="wide")

# --- HELPER FUNCTIONS ---
def clean_col_name(c):
    return str(c).replace('\xa0', '').replace("'", "").strip()

def clean_prefix(name):
    if not isinstance(name, str): return ""
    name = name.replace('\xa0', ' ').strip()
    if len(name) > 2 and name[:2].isupper():
        prefix = name[:2]
        rest = name[2:]
        if rest[:2].upper() == prefix:
            return rest
    return name

def overs_to_balls(overs):
    if pd.isna(overs): return 0
    try:
        o = float(overs)
        complete = int(o)
        balls = round((o - complete) * 10)
        return complete * 6 + balls
    except:
        return 0

# --- MAIN APP ---
st.title("🏏 Proprietary Player Rating Engine")
st.markdown("Upload raw stats, map aliases, and generate the official 10-30 scaled draft board.")

uploaded_file = st.file_uploader("Upload Excel File (.xlsx)", type=["xlsx"])

if uploaded_file:
    xls = pd.ExcelFile(uploaded_file)
    
    bat_sheets = [s for s in xls.sheet_names if 'batting' in s.lower() or 'bat' in s.lower()]
    bowl_sheets = [s for s in xls.sheet_names if 'bowling' in s.lower() or 'bowl' in s.lower()]
    
    # --- STEP 1: EXTRACT RAW DATA ---
    bat_dfs, bowl_dfs = [], []
    for s in bat_sheets:
        header_idx = 1 if 'practice' in s.lower() else 0
        df = pd.read_excel(xls, sheet_name=s, header=header_idx)
        df.columns = [clean_col_name(c) for c in df.columns]
        if 'Player' in df.columns:
            df['Player'] = df['Player'].apply(clean_prefix)
            bat_dfs.append(df)
            
    for s in bowl_sheets:
        header_idx = 1 if 'practice' in s.lower() else 0
        df = pd.read_excel(xls, sheet_name=s, header=header_idx)
        df.columns = [clean_col_name(c) for c in df.columns]
        if 'Player' in df.columns:
            df['Player'] = df['Player'].apply(clean_prefix)
            bowl_dfs.append(df)

    raw_bat = pd.concat(bat_dfs, ignore_index=True) if bat_dfs else pd.DataFrame()
    raw_bowl = pd.concat(bowl_dfs, ignore_index=True) if bowl_dfs else pd.DataFrame()
    
    all_names = set()
    if not raw_bat.empty: all_names.update(raw_bat['Player'].dropna().unique())
    if not raw_bowl.empty: all_names.update(raw_bowl['Player'].dropna().unique())
    
    # --- STEP 2: NAME MERGE EDITOR ---
    st.header("1. Clean Data & Merge Aliases")
    st.write("Edit the **'Merged Name'** column to combine stats for the same player.")
    
    mapping_df = pd.DataFrame({
        "Original Name": sorted(list(all_names)),
        "Merged Name": sorted(list(all_names)) 
    })
    
    edited_mapping = st.data_editor(mapping_df, use_container_width=True, hide_index=True)
    name_map = dict(zip(edited_mapping["Original Name"], edited_mapping["Merged Name"]))

    # --- STEP 3: RATING ENGINE ---
    if st.button("🚀 Calculate Official Rankings", type="primary"):
        with st.spinner("Crunching stats and applying algorithm..."):
            if not raw_bat.empty: raw_bat['Player'] = raw_bat['Player'].map(name_map).fillna(raw_bat['Player'])
            if not raw_bowl.empty: raw_bowl['Player'] = raw_bowl['Player'].map(name_map).fillna(raw_bowl['Player'])
            
            # Aggregate Batting
            for col in ['Runs', 'SR', 'Inns', 'NO']:
                raw_bat[col] = pd.to_numeric(raw_bat[col], errors='coerce').fillna(0)
            raw_bat['Balls_Faced'] = np.where(raw_bat['SR'] > 0, (raw_bat['Runs'] / raw_bat['SR']) * 100, 0)
            raw_bat['Dismissals'] = (raw_bat['Inns'] - raw_bat['NO']).clip(lower=0)
            agg_bat = raw_bat.groupby('Player').agg(
                Inns=('Inns', 'sum'), Runs_bat=('Runs', 'sum'), Balls_Faced=('Balls_Faced', 'sum'), Dismissals=('Dismissals', 'sum')
            ).reset_index()
            
            # Aggregate Bowling
            for col in ['Runs', 'Wkts']:
                raw_bowl[col] = pd.to_numeric(raw_bowl.get(col, 0), errors='coerce').fillna(0)
            raw_bowl['Balls_Bowled'] = raw_bowl['Overs'].apply(overs_to_balls)
            agg_bowl = raw_bowl.groupby('Player').agg(
                Balls_Bowled=('Balls_Bowled', 'sum'), Runs_bowl=('Runs', 'sum'), Wkts=('Wkts', 'sum')
            ).reset_index()
            
            # Process & Smooth
            final_df = pd.merge(agg_bat, agg_bowl, on='Player', how='outer').fillna(0)
            valid = final_df[(final_df['Balls_Faced'] >= 12) | (final_df['Balls_Bowled'] >= 18)].copy()
            
            valid['SR_bat'] = np.where(valid['Balls_Faced'] > 0, (valid['Runs_bat'] / valid['Balls_Faced']) * 100, 0)
            valid['Econ'] = np.where(valid['Balls_Bowled'] > 0, (valid['Runs_bowl'] / valid['Balls_Bowled']) * 6, 999)
            
            mean_avg = valid['Runs_bat'].sum() / (valid['Dismissals'].sum() or 1)
            mean_sr = (valid['Runs_bat'].sum() / valid['Balls_Faced'].sum()) * 100
            mean_econ = (valid['Runs_bowl'].sum() / valid['Balls_Bowled'].sum()) * 6
            
            valid['Sm_Avg'] = (valid['Runs_bat'] + (mean_avg * 3)) / (valid['Dismissals'] + 3)
            valid['Sm_SR'] = ((valid['Runs_bat'] + (mean_sr/100 * 30)) / (valid['Balls_Faced'] + 30)) * 100
            valid['Sm_Econ'] = ((valid['Runs_bowl'] + (mean_econ/6 * 30)) / (valid['Balls_Bowled'] + 30)) * 6
            valid['Sm_Avg_bowl'] = (valid['Runs_bowl'] + (mean_econ * 5)) / (valid['Wkts'] + 5)
            
            z_runs = (valid['Runs_bat'] - valid['Runs_bat'].mean()) / valid['Runs_bat'].std()
            z_avg = (valid['Sm_Avg'] - valid['Sm_Avg'].mean()) / valid['Sm_Avg'].std()
            z_sr = (valid['Sm_SR'] - valid['Sm_SR'].mean()) / valid['Sm_SR'].std()
            valid['Bat_Score'] = z_runs*0.3 + z_avg*0.4 + z_sr*0.3
            
            z_wkts = (valid['Wkts'] - valid['Wkts'].mean()) / valid['Wkts'].std()
            z_econ = (valid['Sm_Econ'].mean() - valid['Sm_Econ']) / valid['Sm_Econ'].std()
            z_avg_bowl = (valid['Sm_Avg_bowl'].mean() - valid['Sm_Avg_bowl']) / valid['Sm_Avg_bowl'].std()
            valid['Bowl_Score'] = z_wkts*0.4 + z_econ*0.35 + z_avg_bowl*0.25
            
            def assign_role(row):
                if row['Balls_Faced'] >= 15 and row['Balls_Bowled'] >= 18: return 'All-Rounder'
                elif row['Balls_Bowled'] >= 18: return 'Bowler'
                else: return 'Batter'
            valid['Role'] = valid.apply(assign_role, axis=1)
            
            def calc_final(row):
                if row['Role'] == 'Batter': return row['Bat_Score']
                elif row['Role'] == 'Bowler': return row['Bowl_Score']
                else: return (row['Bat_Score'] * 0.5 + row['Bowl_Score'] * 0.5) * 1.3
            valid['Final_Raw'] = valid.apply(calc_final, axis=1)
            
            min_raw, max_raw = np.percentile(valid['Final_Raw'], 1), np.percentile(valid['Final_Raw'], 99)
            valid['Rating'] = ((valid['Final_Raw'] - min_raw) / (max_raw - min_raw)) * 20.0 + 10.0
            valid['Rating'] = valid['Rating'].clip(lower=10.0, upper=30.0).round(1)
            
            def format_stats(row):
                if row['Role'] == 'Batter': return f"{int(row['Runs_bat'])} Runs (SR: {row['SR_bat']:.1f})"
                elif row['Role'] == 'Bowler': return f"{int(row['Wkts'])} Wkts (Econ: {row['Econ']:.1f})"
                else: return f"{int(row['Runs_bat'])} Runs, {int(row['Wkts'])} Wkts"
            valid['Key Stats'] = valid.apply(format_stats, axis=1)
            
            output = valid[['Player', 'Role', 'Key Stats', 'Rating']].sort_values('Rating', ascending=False).reset_index(drop=True)
            output.index = output.index + 1
            
            # Save into session state to display filtering options
            st.session_state['draft_board'] = output
            
    # --- STEP 4: INTERACTIVE DRAFT BOARD ---
    if 'draft_board' in st.session_state:
        st.header("🏆 Interactive Draft Board")
        
        col1, col2 = st.columns(2)
        with col1:
            search_query = st.text_input("🔍 Search Player")
        with col2:
            role_filter = st.selectbox("🎯 Filter by Role", ["All", "Batter", "Bowler", "All-Rounder"])
            
        display_df = st.session_state['draft_board']
        
        if search_query:
            display_df = display_df[display_df['Player'].str.contains(search_query, case=False, na=False)]
        if role_filter != "All":
            display_df = display_df[display_df['Role'] == role_filter]
            
        st.dataframe(display_df, use_container_width=True)
        
        csv = display_df.to_csv(index=True).encode('utf-8')
        st.download_button(label="📥 Download Draft Cheatsheet (CSV)", data=csv, file_name='cricket_draft_cheatsheet.csv', mime='text/csv')
