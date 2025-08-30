import pandas as pd


def generate_claim_reports(df, df_claims_data_param, bool_hub):
    # --- Base copies ---
    df_claims = df.copy()
    df_claims_data = df_claims_data_param.copy()

    # --- Clean company names early ---
    df_claims["COMPANYNAME"] = (
        df_claims["COMPANYNAME"]
        .astype("string")
        .str.replace(".", "", regex=False)
        .str.strip()
    )

    # --- Clean claims data Policy Number ---
    df_claims_data["Policy Number"] = df_claims_data["Policy Number"].str.replace(
        "#", "", regex=False
    )
    df_claims_data = df_claims_data.drop_duplicates(subset=["Policy Number"])

    # --- Merge on Policy Number ---
    df_claims = df_claims.merge(
        df_claims_data,
        left_on="TXT_POLICY_NO_CHAR",
        right_on="Policy Number",
        how="left",
    )

    # --- Date conversions ---
    df_claims["Policy From"] = pd.to_datetime(df_claims["Policy From"], format="mixed")
    df_claims["Policy Upto"] = pd.to_datetime(df_claims["Policy Upto"], format="mixed")
    df_claims["DAT_LOSS_DATE"] = pd.to_datetime(
        df_claims["DAT_LOSS_DATE"], format="mixed"
    )
    df_claims["DAT_ACCOUNTING_DATE"] = pd.to_datetime(
        df_claims["DAT_ACCOUNTING_DATE"], format="mixed"
    )

    # --- Hub filter ---
    cutoff = pd.Timestamp("2023-04-01")
    mask = (
        df_claims["Policy From"] < cutoff
        if bool_hub
        else df_claims["Policy From"] >= cutoff
    )
    df_claims = df_claims.loc[mask].copy()

    # --- Derived fields ---
    df_claims["Total claim amount"] = (
        df_claims["CUR_DEBIT_BALANCE"] - df_claims["CUR_CREDIT_BALANCE"]
    )
    df_claims["Origin Office"] = df_claims["TXT_POLICY_NO_CHAR"].str[:6]

    # --- Mapping (single source of truth) ---
    columns_map = {
        "Origin Office": "Origin Office",
        "TXT_UIIC_OFF_CODE": "UIIC Office Code",
        "COMPANYNAME": "Name of coinsurer",
        "TXT_LEADER_OFFICE_CODE": "Follower Office Code",
        "Customer Name": "Name of insured",
        "TXT_DEPARTMENTNAME": "Department",
        "TXT_POLICY_NO_CHAR": "Policy Number",
        "Policy From": "Policy start date",
        "Policy Upto": "Policy end date",
        "NUM_SHARE_PCT": "Percentage of share",
        "TXT_MASTER_CLAIM_NO": "Claim Number",
        "DAT_LOSS_DATE": "Date of loss",
        "TXT_NATURE_OF_LOSS": "Cause of loss",
        "Total claim amount": "Total claim amount",
        "NUM_VOUCHER_NO": "Settlement voucher number",
        "DAT_ACCOUNTING_DATE": "Voucher date",
        "TXT_URN_CODE": "URN Code",
    }

    # --- Restrict to needed cols + rename ---
    df_claims = df_claims[list(columns_map.keys())].rename(columns=columns_map)

    # --- Post-formatting ---
    df_claims["Settlement voucher number"] = "'" + df_claims[
        "Settlement voucher number"
    ].astype(str)

    return df_claims
