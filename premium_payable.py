import pandas as pd


def generate_premium_payable(df, bool_hub):
    df_pp = df.copy()

    # filter hub/non-hub
    df_pp["DAT_POLICY_START_DATE"] = pd.to_datetime(
        df_pp["DAT_POLICY_START_DATE"], format="mixed"
    )
    cutoff = pd.Timestamp("2023-04-01")
    mask = (
        df_pp["DAT_POLICY_START_DATE"] < cutoff
        if bool_hub
        else df_pp["DAT_POLICY_START_DATE"] >= cutoff
    )
    df_pp = df_pp.loc[mask].copy()

    # clean company names
    df_pp["COMPANYNAME"] = (
        df_pp["COMPANYNAME"]
        .astype("string")
        .str.replace(".", "", regex=False)
        .str.strip()
        .astype("category")
    )

    # compute derived fields
    df_pp = df_pp.assign(
        GST_on_admin_charges=df_pp["CUR_ADMIN_CHARGE"] * 0.18,
        NUM_VOUCHER_NO="'" + df_pp["NUM_VOUCHER_NO"].astype(str),
        DAT_POLICY_END_DATE=pd.to_datetime(
            df_pp["DAT_POLICY_END_DATE"], format="mixed"
        ),
        DAT_ACCOUNTING_DATE=pd.to_datetime(
            df_pp["DAT_ACCOUNTING_DATE"], format="mixed"
        ),
        Net_Premium_payable=(
            df_pp["PREMIUM"]
            - df_pp["CUR_ADMIN_CHARGE"]
            - df_pp["NUM_TPA_SER_CHARGE_SHARE"]
            - df_pp["NUM_COMMISSION_SHARE"]
            - (df_pp["CUR_ADMIN_CHARGE"] * 0.18)
        ),
        Origin_Office=df_pp["TXT_POLICY_NO_CHAR"].str[:6],
    )

    # rename columns
    columns_map = {
        "Origin_Office": "Origin Office",
        "TXT_UIIC_OFFICE_CD": "UIIC Office Code",
        "COMPANYNAME": "Name of coinsurer",
        "TXT_FOLLOWER_OFF_CD": "Follower Office Code",
        "TXT_NAME_OF_INSURED": "Name of insured",
        "TXT_DEPARTMENTNAME": "Department",
        "TXT_POLICY_NO_CHAR": "Policy Number",
        "NUM_ENDT_NO": "Endorsement number",
        "DAT_POLICY_START_DATE": "Policy start date",
        "DAT_POLICY_END_DATE": "Policy end date",
        "DAT_ACCOUNTING_DATE": "Accounting date",
        "NUM_VOUCHER_NO": "Voucher number",
        "NUM_SHARE_PCT": "Percentage of share",
        "CUR_SUM_INSURED": "Current sum insured",
        "TXT_URN_CODE": "URN Code",
        "PREMIUM": "Premium",
        "NUM_COMMISSION_SHARE": "Total Brokerage amount",
        "NUM_TPA_SER_CHARGE_SHARE": "TPA Service Charges amount",
        "CUR_ADMIN_CHARGE": "Admin charges",
        "GST_on_admin_charges": "GST on admin charges",
        "Net_Premium_payable": "Net Premium payable",
    }
    df_pp = df_pp[list(columns_map.keys())]
    return df_pp.rename(columns=columns_map)
