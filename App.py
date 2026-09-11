    mean_raw = valid['Final_Raw'].mean()
    std_raw = valid['Final_Raw'].std() or 1
    
    valid['AI Rating'] = 20.0 + ((valid['Final_Raw'] - mean_raw) / std_raw) * 3.33
    valid['AI Rating'] = valid['AI Rating'].clip(lower=10.0, upper=30.0).round(1)
    
    for col, new_col in [('Bat_Score', 'Bat_Rating'), ('Bowl_Score', 'Bowl_Rating'), ('Fielding_Score', 'Field_Rating')]:
        mean_val = valid[col].mean()
        std_val = valid[col].std() or 1
        valid[new_col] = 20.0 + ((valid[col] - mean_val) / std_val) * 3.33
        valid[new_col] = valid[new_col].clip(lower=10.0, upper=30.0).round(1)

    valid['Tier'] = pd.cut(valid['AI Rating'], bins=[-1, 16.9, 22.9, 26.9, 31], labels=["Bronze", "Silver", "Gold", "Platinum"])
    return valid
