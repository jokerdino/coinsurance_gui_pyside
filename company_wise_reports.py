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
    """
    Merge Premium Payable and Claim Receivable reports by company.
    Generates company-specific Excel files (CR, PP, or both) and pivot tables.

    Args:
        df_pp: Premium Payable dataframe
        df_claims_reports: Claims Receivable dataframe
        df_claims_data: Claims data dataframe
        path_string_wip: Output folder path
        bool_hub: Hub flag (True = hub entries, False = non-hub)
    """

    os.makedirs(path_string_wip, exist_ok=True)

    # --- Process inputs ---
    df_claims_receivable = generate_claim_reports(
        df_claims_reports, df_claims_data, bool_hub
    )
    df_premium_payable = generate_premium_payable(df_pp, bool_hub)

    claims_companies = set(df_claims_receivable["Name of coinsurer"].unique())
    premium_companies = set(df_premium_payable["Name of coinsurer"].unique())

    # classify companies
    companies_claims_only = sorted(claims_companies - premium_companies)
    companies_premium_only = sorted(premium_companies - claims_companies)
    companies_both = sorted(claims_companies & premium_companies)

    # --- Helper for saving ---
    def save_company_file(company, df_claims=None, df_premium=None):
        company_dir = os.path.join(path_string_wip, company)
        os.makedirs(company_dir, exist_ok=True)

        file_path = os.path.join(company_dir, f"{company}.xlsx")

        with pd.ExcelWriter(file_path, datetime_format="dd/mm/yyyy") as writer:
            if df_claims is not None:
                df_claims.to_excel(writer, index=False, sheet_name="CR")
                worksheet_formatter(writer, "CR")

            if df_premium is not None:
                df_premium.sort_values(
                    [
                        "Policy Number",
                        "Endorsement number",
                        "Accounting date",
                        "Voucher number",
                    ],
                    ascending=[True, True, True, True],
                ).to_excel(writer, index=False, sheet_name="PP")
                worksheet_formatter(writer, "PP")

        generate_pivot_table(file_path, company, path_string_wip)

    # --- Write outputs ---
    for company in companies_claims_only:
        df_claims = df_claims_receivable[
            df_claims_receivable["Name of coinsurer"].eq(company)
        ]
        save_company_file(company, df_claims=df_claims)

    for company in companies_premium_only:
        df_premium = df_premium_payable[
            df_premium_payable["Name of coinsurer"].eq(company)
        ]
        save_company_file(company, df_premium=df_premium)

    for company in companies_both:
        df_claims = df_claims_receivable[
            df_claims_receivable["Name of coinsurer"].eq(company)
        ]
        df_premium = df_premium_payable[
            df_premium_payable["Name of coinsurer"].eq(company)
        ]
        save_company_file(company, df_claims=df_claims, df_premium=df_premium)
