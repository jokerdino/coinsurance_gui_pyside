import os

import pandas as pd

from claims_receivable import generate_claim_reports
from premium_payable import generate_premium_payable
from pivot_table import generate_pivot_table


def worksheet_formatter(writer, sheet_name):
    format_workbook = writer.book
    format_currency = format_workbook.add_format({"num_format": "##,##,#0"})

    format_worksheet = writer.sheets[sheet_name]

    if sheet_name == "PP":
        format_worksheet.set_column("P:U", 11, format_currency)
    elif sheet_name == "PR":
        format_worksheet.set_column("Q:V", 11, format_currency)
    elif sheet_name == "CR":
        format_worksheet.set_column("N:N", 11, format_currency)
    format_worksheet.autofit()


def merge_files(df_pp, df_claims_reports, df_claims_data, path_string_wip, bool_hub):
    if not os.path.exists(path_string_wip):
        os.mkdir(path_string_wip)
    df_claims_receivable = generate_claim_reports(
        df_claims_reports, df_claims_data, bool_hub
    )
    df_premium_payable = generate_premium_payable(df_pp, bool_hub)

    list_claims_coinsurers = df_claims_receivable["Name of coinsurer"].unique().tolist()
    list_premium_payable_coinsurers = (
        df_premium_payable["Name of coinsurer"].unique().tolist()
    )

    list_concatenated = list(
        set(list_claims_coinsurers).intersection(list_premium_payable_coinsurers)
    )
    list_claims_alone = list(
        set(list_claims_coinsurers) - set(list_premium_payable_coinsurers)
    )
    list_premium_alone = list(
        set(list_premium_payable_coinsurers) - set(list_claims_coinsurers)
    )

    # find companies only in claims list and premium list and write them to file
    # also run pivot table for them
    for company in sorted(list_claims_alone):
        if not os.path.exists(f"{path_string_wip}/{company}"):
            os.mkdir(f"{path_string_wip}/{company}")
        df_company_claims = df_claims_receivable[
            df_claims_receivable["Name of coinsurer"].str.contains(company)
        ]

        with pd.ExcelWriter(
            f"{path_string_wip}/{company}/{company}.xlsx", datetime_format="dd/mm/yyyy"
        ) as writer:
            df_company_claims.to_excel(writer, index=False, sheet_name="CR")
            worksheet_formatter(writer, "CR")

        generate_pivot_table(
            f"{path_string_wip}/{company}/{company}.xlsx", company, path_string_wip
        )

    for company in sorted(list_premium_alone):
        if not os.path.exists(f"{path_string_wip}/{company}"):
            os.mkdir(f"{path_string_wip}/{company}")
        df_company_premium = df_premium_payable[
            df_premium_payable["Name of coinsurer"].str.contains(company)
        ]

        with pd.ExcelWriter(
            f"{path_string_wip}/{company}/{company}.xlsx", datetime_format="dd/mm/yyyy"
        ) as writer:
            df_company_premium.sort_values(
                [
                    "Policy Number",
                    "Endorsement number",
                    "Accounting date",
                    "Voucher number",
                ],
                ascending=[True, True, True, True],
            ).to_excel(writer, index=False, sheet_name="PP")
            worksheet_formatter(writer, "PP")

        generate_pivot_table(
            f"{path_string_wip}/{company}/{company}.xlsx", company, path_string_wip
        )

    for company in sorted(list_concatenated):
        if not os.path.exists(f"{path_string_wip}/{company}"):
            os.mkdir(f"{path_string_wip}/{company}")

        df_company_claims = df_claims_receivable[
            df_claims_receivable["Name of coinsurer"].str.contains(company)
        ]
        df_company_premium = df_premium_payable[
            df_premium_payable["Name of coinsurer"].str.contains(company)
        ]

        with pd.ExcelWriter(
            f"{path_string_wip}/{company}/{company}.xlsx", datetime_format="dd/mm/yyyy"
        ) as writer:
            df_company_claims.to_excel(writer, index=False, sheet_name="CR")
            df_company_premium.sort_values(
                [
                    "Policy Number",
                    "Endorsement number",
                    "Accounting date",
                    "Voucher number",
                ],
                ascending=[True, True, True, True],
            ).to_excel(writer, index=False, sheet_name="PP")

            worksheet_formatter(writer, "PP")
            worksheet_formatter(writer, "CR")

        generate_pivot_table(
            f"{path_string_wip}/{company}/{company}.xlsx", company, path_string_wip
        )
