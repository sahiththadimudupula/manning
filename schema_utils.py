import pandas as pd


# -------------------------------
# Normalize column names
# -------------------------------
def normalize_columns(df: pd.DataFrame) -> pd.DataFrame:
    df = df.copy()

    df.columns = (
        df.columns
        .astype(str)
        .str.strip()
        .str.replace(r"\s+", "_", regex=True)
    )

    # --- Fix known inconsistencies ---
    df = df.rename(columns={
        "HO_Scientific_Manpower": "BE_Scientific_Manpower",
        "HO_Final_Manpower": "BE_Final_Manpower",
        "N_shift": "N_shifts",
        "N_Shift": "N_shifts",
        "N_Shifts": "N_shifts",
    })

    return df


# -------------------------------
# Compute N_shifts
# -------------------------------
def compute_n_shifts(df: pd.DataFrame) -> pd.DataFrame:
    def calc(row):
        g = float(row.get("General_Shift", 0) or 0)
        a = float(row.get("Shift_A", 0) or 0)
        b = float(row.get("Shift_B", 0) or 0)
        c = float(row.get("Shift_C", 0) or 0)

        # Only General shift
        if a == 0 and b == 0 and c == 0 and g > 0:
            return 1

        # Count active A/B/C
        active = sum([a > 0, b > 0, c > 0])

        # No General shift
        if g == 0 and active > 0:
            return active

        # Fallback
        return max(1, active)

    df["N_shifts"] = df.apply(calc, axis=1)
    return df


# -------------------------------
# Main function
# -------------------------------
def align_and_validate_schemas(spinning_path, wtt_path, rugs_path):

    # --- Load Excel ---
    spin_df = pd.read_excel(spinning_path, sheet_name="Spinning")
    wtt_df = pd.read_excel(wtt_path, sheet_name="WTT")
    rugs_df = pd.read_excel(rugs_path, sheet_name="Rugs")

    # --- Normalize columns ---
    spin_df = normalize_columns(spin_df)
    wtt_df = normalize_columns(wtt_df)
    rugs_df = normalize_columns(rugs_df)

    # --- Ensure Dept column ---
    for df in [spin_df, wtt_df, rugs_df]:
        if "Dept_Machine_Name" not in df.columns and "Department" in df.columns:
            df["Dept_Machine_Name"] = df["Department"]

    # --- Enforce Business column (DO NOT trust Excel) ---
    spin_df["Business"] = "Spinning"
    wtt_df["Business"] = "WTT"
    rugs_df["Business"] = "Rugs"

    # --- Ensure shift columns exist ---
    shift_cols = ["General_Shift", "Shift_A", "Shift_B", "Shift_C"]
    for df in [spin_df, wtt_df, rugs_df]:
        for col in shift_cols:
            if col not in df.columns:
                df[col] = 0

    # --- Ensure numeric columns exist ---
    for df in [spin_df, wtt_df, rugs_df]:
        for col in ["Contractors", "Company_Associate"]:
            if col not in df.columns:
                df[col] = 0

    # --- Clean strings (critical for filters) ---
    for df in [spin_df, wtt_df, rugs_df]:
        df["Business"] = df["Business"].astype(str).str.strip()
        if "Section" in df.columns:
            df["Section"] = df["Section"].astype(str).str.strip()
        else:
            df["Section"] = ""

    # --- Compute N_shifts ---
    spin_df = compute_n_shifts(spin_df)
    wtt_df = compute_n_shifts(wtt_df)
    rugs_df = compute_n_shifts(rugs_df)

    # --- Use Rugs as base schema ---
    base_columns = list(rugs_df.columns)

    def align_to_base(df: pd.DataFrame) -> pd.DataFrame:
        df = df.copy()

        # Add missing columns
        for col in base_columns:
            if col not in df.columns:
                df[col] = 0

        # Drop extra columns
        extra_cols = set(df.columns) - set(base_columns)
        if extra_cols:
            df = df.drop(columns=list(extra_cols))

        # Reorder exactly like Rugs
        df = df[base_columns]

        return df

    spin_df = align_to_base(spin_df)
    wtt_df = align_to_base(wtt_df)
    rugs_df = align_to_base(rugs_df)

    return spin_df, wtt_df, rugs_df