import pandas as pd


def generate_claim_reports(df, df_claims_data_param, bool_hub):
    df_claims_reports = df.copy()
    df_claims_data = df_claims_data_param.copy()
    df_claims_reports["COMPANYNAME"] = df_claims_reports["COMPANYNAME"].astype(
        "category"
    )

    df_claims_data["Policy Number"] = df_claims_data["Policy Number"].str.replace(
        "#", ""
    )

    df_claims_data = df_claims_data.drop_duplicates(subset=["Policy Number"])
    # rearranging the columns to our preference

    df_claims = df_claims_reports.merge(
        df_claims_data,
        left_on=("TXT_POLICY_NO_CHAR"),
        right_on=("Policy Number"),
        how="left",
    )
    df_claims["Policy From"] = pd.to_datetime(df_claims["Policy From"], format="mixed")
    if bool_hub:
        df_claims = df_claims[df_claims["Policy From"] < "2023-04-01"]
    else:
        df_claims = df_claims[df_claims["Policy From"] > "2023-03-31"]

    df_claims["Policy Upto"] = pd.to_datetime(df_claims["Policy Upto"], format="mixed")

    df_claims["Total claim amount"] = (
        df_claims["CUR_DEBIT_BALANCE"] - df_claims["CUR_CREDIT_BALANCE"]
    )

    df_claims["DAT_LOSS_DATE"] = pd.to_datetime(
        df_claims["DAT_LOSS_DATE"], format="mixed"
    )
    df_claims["DAT_ACCOUNTING_DATE"] = pd.to_datetime(
        df_claims["DAT_ACCOUNTING_DATE"], format="mixed"
    )

    df_claims["COMPANYNAME"] = df_claims["COMPANYNAME"].str.replace(".", "")
    df_claims["COMPANYNAME"] = df_claims["COMPANYNAME"].str.rstrip()
    df_claims["Origin Office"] = df_claims["TXT_POLICY_NO_CHAR"].str[:6]
    df_claims = df_claims[
        [
            "Origin Office",
            "TXT_UIIC_OFF_CODE",
            "COMPANYNAME",
            "TXT_LEADER_OFFICE_CODE",
            "Customer Name",
            "TXT_DEPARTMENTNAME",
            "TXT_POLICY_NO_CHAR",
            "Policy From",
            "Policy Upto",
            "NUM_SHARE_PCT",
            "TXT_MASTER_CLAIM_NO",
            "DAT_LOSS_DATE",
            "TXT_NATURE_OF_LOSS",
            "Total claim amount",
            "NUM_VOUCHER_NO",
            "DAT_ACCOUNTING_DATE",
            "TXT_URN_CODE",
        ]
    ]

    df_claims = df_claims.set_axis(
        [
            "Origin Office",
            "UIIC Office Code",
            "Name of coinsurer",
            "Follower Office Code",
            "Name of insured",
            "Department",
            "Policy Number",
            "Policy start date",
            "Policy end date",
            "Percentage of share",
            "Claim Number",
            "Date of loss",
            "Cause of loss",
            "Total claim amount",
            "Settlement voucher number",
            "Voucher date",
            "URN Code",
        ],
        axis=1,
        copy=False,
    )

    df_claims["Settlement voucher number"] = "'" + df_claims[
        "Settlement voucher number"
    ].astype(str)

    return df_claims
