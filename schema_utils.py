import pandas as pd

def align_and_validate_schemas(spinning_path, wtt_path):
    spin_df = pd.read_excel(spinning_path, sheet_name="Spinning")
    wtt_df = pd.read_excel(wtt_path, sheet_name="WTT")

    # normalize mismatched columns
    wtt_df = wtt_df.rename(columns={
        "HO_Scientific_Manpower": "BE_Scientific_Manpower",
        "HO_Final_Manpower": "BE_Final_Manpower",
    })

    for df in [spin_df, wtt_df]:
        if "Dept_Machine_Name" not in df.columns and "Department" in df.columns:
            df["Dept_Machine_Name"] = df["Department"]

    # schema validation
    spin_cols = set(spin_df.columns)
    wtt_cols = set(wtt_df.columns)

    missing_in_wtt = spin_cols - wtt_cols
    missing_in_spin = wtt_cols - spin_cols

    if missing_in_wtt or missing_in_spin:
        raise ValueError(
            f"Schema mismatch\n"
            f"Missing in WTT: {missing_in_wtt}\n"
            f"Missing in Spinning: {missing_in_spin}"
        )

    # align order
    wtt_df = wtt_df[spin_df.columns]

    return spin_df, wtt_df