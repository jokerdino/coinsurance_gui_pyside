import pandas as pd
from company_wise_reports import worksheet_formatter


def split_files(excel_filename, follower_office_code):
    bool_claims, bool_premium = False, False

    try:
        df_claims = pd.read_excel(
            excel_filename,
            "CR",
            converters={"Follower Office Code": str, "Origin Office": str},
        )
        df_claims_filtered = df_claims[
            df_claims["Follower Office Code"] == follower_office_code
        ]
        if len(df_claims_filtered) > 0:
            bool_claims = True
    except ValueError:
        pass

    try:
        df_premium = pd.read_excel(
            excel_filename,
            "PP",
            converters={"Follower Office Code": str, "Origin Office": str},
        )
        df_premium_filtered = df_premium[
            df_premium["Follower Office Code"] == follower_office_code
        ]
        if len(df_premium_filtered) > 0:
            bool_premium = True
    except ValueError:
        pass

    company_name = excel_filename.split(".", 1)[0]

    if bool_claims and bool_premium:
        with pd.ExcelWriter(
            f"{company_name}_{follower_office_code}.xlsx", datetime_format="dd/mm/yyyy"
        ) as writer:
            df_claims_filtered.to_excel(writer, index=False, sheet_name="CR")
            df_premium_filtered.to_excel(writer, index=False, sheet_name="PP")

            worksheet_formatter(writer, "PP")
            worksheet_formatter(writer, "CR")

    elif bool_claims:
        with pd.ExcelWriter(
            f"{company_name}_{follower_office_code}.xlsx", datetime_format="dd/mm/yyyy"
        ) as writer:
            df_claims_filtered.to_excel(writer, index=False, sheet_name="CR")
            worksheet_formatter(writer, "CR")

    elif bool_premium:
        with pd.ExcelWriter(
            f"{company_name}_{follower_office_code}.xlsx", datetime_format="dd/mm/yyyy"
        ) as writer:
            df_premium_filtered.to_excel(writer, index=False, sheet_name="PP")
            worksheet_formatter(writer, "PP")
    else:
        pass
