import pandas as pd


def generate_premium_payable(df, bool_hub):
    df_pp_input = df.copy()
    df_pp_input["DAT_POLICY_START_DATE"] = pd.to_datetime(
        df_pp_input["DAT_POLICY_START_DATE"], format="mixed"
    )
    if bool_hub:
        df_pp = df_pp_input[df_pp_input["DAT_POLICY_START_DATE"] < "2023-04-01"].copy()
    else:
        df_pp = df_pp_input[df_pp_input["DAT_POLICY_START_DATE"] > "2023-03-31"].copy()

    df_pp["GST on admin charges"] = df_pp["CUR_ADMIN_CHARGE"] * 18 / 100
    df_pp["NUM_VOUCHER_NO"] = "'" + df_pp["NUM_VOUCHER_NO"].astype(str)

    df_pp["DAT_POLICY_END_DATE"] = pd.to_datetime(
        df_pp["DAT_POLICY_END_DATE"], format="mixed"
    )

    df_pp["DAT_ACCOUNTING_DATE"] = pd.to_datetime(
        df_pp["DAT_ACCOUNTING_DATE"], format="mixed"
    )
    df_pp["COMPANYNAME"] = df_pp["COMPANYNAME"].astype("category")

    df_pp["COMPANYNAME"] = df_pp["COMPANYNAME"].str.replace(".", "")
    df_pp["COMPANYNAME"] = df_pp["COMPANYNAME"].str.rstrip()

    df_pp["Net Premium payable"] = (
        df_pp["PREMIUM"]
        - df_pp["CUR_ADMIN_CHARGE"]
        - df_pp["NUM_TPA_SER_CHARGE_SHARE"]
        - df_pp["NUM_COMMISSION_SHARE"]
        - df_pp["GST on admin charges"]
    )

    df_pp["Origin Office"] = df_pp["TXT_POLICY_NO_CHAR"].str[:6]
    df_pp = df_pp[
        [
            "Origin Office",
            "TXT_UIIC_OFFICE_CD",
            "COMPANYNAME",
            "TXT_FOLLOWER_OFF_CD",
            "TXT_NAME_OF_INSURED",
            "TXT_DEPARTMENTNAME",
            "TXT_POLICY_NO_CHAR",
            "NUM_ENDT_NO",
            "DAT_POLICY_START_DATE",
            "DAT_POLICY_END_DATE",
            "DAT_ACCOUNTING_DATE",
            "NUM_VOUCHER_NO",
            "NUM_SHARE_PCT",
            "CUR_SUM_INSURED",
            "TXT_URN_CODE",
            "PREMIUM",
            "NUM_COMMISSION_SHARE",
            "NUM_TPA_SER_CHARGE_SHARE",
            "CUR_ADMIN_CHARGE",
            "GST on admin charges",
            "Net Premium payable",
        ]
    ]
    df_pp = df_pp.set_axis(
        [
            "Origin Office",
            "UIIC Office Code",
            "Name of coinsurer",
            "Follower Office Code",
            "Name of insured",
            "Department",
            "Policy Number",
            "Endorsement number",
            "Policy start date",
            "Policy end date",
            "Accounting date",
            "Voucher number",
            "Percentage of share",
            "Current sum insured",
            "URN Code",
            "Premium",
            "Total Brokerage amount",
            "TPA Service Charges amount",
            "Admin charges",
            "GST on admin charges",
            "Net Premium payable",
        ],
        axis=1,
        copy=False,
    )
    return df_pp
