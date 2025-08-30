import os

import pandas as pd
import numpy as np

from company_wise_reports import worksheet_formatter


def premium_receivable_function(df_premium_receivable, folder_name):
    if not os.path.exists(f"{folder_name}_PR"):
        os.mkdir(f"{folder_name}_PR")

    df_premium_receivable["COMPANYNAME"] = df_premium_receivable["COMPANYNAME"].astype(
        "category"
    )

    df_premium_receivable["COMPANYNAME"] = df_premium_receivable[
        "COMPANYNAME"
    ].str.replace(".", "")
    df_premium_receivable["COMPANYNAME"] = df_premium_receivable[
        "COMPANYNAME"
    ].str.rstrip()

    df_premium_receivable["DAT_ACCOUNTING_DATE"] = pd.to_datetime(
        df_premium_receivable["DAT_ACCOUNTING_DATE"], format="mixed"
    )

    df_premium_receivable.drop(
        df_premium_receivable.filter(regex="Aggrega").columns, axis=1, inplace=True
    )
    df_premium_receivable["NUM_VOUCHER_NO"] = "'" + df_premium_receivable[
        "NUM_VOUCHER_NO"
    ].astype(str)

    df_premium_receivable["DAT_POLICY_START_DATE"] = pd.to_datetime(
        df_premium_receivable["DAT_POLICY_START_DATE"], format="mixed"
    )
    df_premium_receivable["DAT_POLICY_END_DATE"] = pd.to_datetime(
        df_premium_receivable["DAT_POLICY_END_DATE"], format="mixed"
    )
    df_premium_receivable["DAT_ACCOUNTING_DATE"] = pd.to_datetime(
        df_premium_receivable["DAT_ACCOUNTING_DATE"], format="mixed"
    )
    # TOTAL 40 COLUMNS: REMOVING 22 COLUMNS REMAINING 18 COLUMNS
    df_premium_receivable["Origin Office"] = df_premium_receivable[
        "TXT_POLICY_NO_CHAR"
    ].str[:6]
    df_premium_receivable = df_premium_receivable[
        [
            "COMPANYNAME",
            "TXT_LEADER_OFFICE_CODE",
            "Origin Office",
            "TXT_UIIC_OFF_CODE",
            "TXT_LEADER_POLICY_NO",
            "TXT_POLICY_NO_CHAR",
            "NUM_ENDT_NO",
            "TXT_URN_CODE",
            "TXT_DEPARTMENTNAME",
            "TXT_NAME_OF_INSURED",
            "DAT_POLICY_START_DATE",
            "DAT_POLICY_END_DATE",
            "DAT_ACCOUNTING_DATE",
            "CUR_SUM_INSURED",
            "NUM_SHARE_PCT",
            "NUM_VOUCHER_NO",
            "PREMIUM",
            "NUM_COMMISSION_SHARE",
            "NUM_TPA_SER_CHARGE_SHARE",
        ]
    ]

    df_premium_receivable = df_premium_receivable.set_axis(
        [
            "Name of coinsurer",
            "Leader Office Code",
            "Origin Office",
            "UIIC Office Code",
            "Leader Policy Number",
            "United India Policy Number",
            "Endorsement number",
            "URN number",
            "Department",
            "Name of insured",
            "Policy start date",
            "Policy end date",
            "Accounting date",
            "Sum insured",
            "Percentage share (%)",
            "Voucher number",
            "Premium",
            "Brokerage amount",
            "TPA Service charges",
        ],
        axis=1,
        copy=False,
    )  # ,'Admin charges (','','','','','','','','',''])

    df_premium_receivable["Admin charges"] = df_premium_receivable["Premium"] * 1 / 100
    df_premium_receivable["GST on Admin charges"] = (
        df_premium_receivable["Admin charges"] * 18 / 100
    )
    df_premium_receivable["Net Premium receivable"] = (
        df_premium_receivable["Premium"]
        - df_premium_receivable["Brokerage amount"]
        - df_premium_receivable["TPA Service charges"]
        - df_premium_receivable["Admin charges"]
        - df_premium_receivable["GST on Admin charges"]
    )

    df_premium_receivable = df_premium_receivable[
        df_premium_receivable["Policy start date"] > "2023-03-31"
    ]
    # df_premium_receivable.to_excel("Merged.xlsx", index=False)

    unique_coinsurer_name = df_premium_receivable["Name of coinsurer"].unique()

    for i in unique_coinsurer_name:
        df_coinsurer = df_premium_receivable[
            df_premium_receivable["Name of coinsurer"].str.contains(i)
        ]
        pivot_coinsurer = pd.pivot_table(
            df_coinsurer,
            index=["Origin Office", "Leader Office Code"],
            values="Net Premium receivable",
            aggfunc="sum",
        )
        pivot_coinsurer.reset_index(inplace=True)
        pivot_coinsurer.loc["Total"] = pivot_coinsurer.sum(numeric_only=True, axis=0)
        pivot_coinsurer.style.format({"Net Premium receivable": "{0:,.2f}"})
        with pd.ExcelWriter(
            f"{folder_name}_PR/{i}_Premium_receivable.xlsx",
            datetime_format="dd/mm/yyyy",
        ) as writer:
            df_coinsurer.to_excel(writer, sheet_name="PR", index=False)
            pivot_coinsurer.to_excel(writer, sheet_name="Summary", index=False)
            worksheet_formatter(writer, "PR")

            format_workbook = writer.book
            format_currency = format_workbook.add_format({"num_format": "##,##,#0"})
            format_bold = format_workbook.add_format(
                {"num_format": "##,##,#", "bold": True}
            )
            format_header = format_workbook.add_format(
                {"bold": True, "text_wrap": True, "valign": "top"}
            )
            format_worksheet = writer.sheets["Summary"]

            format_worksheet.set_column("C:C", 12, format_bold)
            format_worksheet.set_column("C:C", 12, format_currency)
            format_worksheet.set_row(-1, 12, format_bold)
            format_worksheet.set_row(0, None, format_header)


# df_premium_receivable = pd.read_csv('premium_receivable.csv', converters={'TXT_LEADER_OFFICE_CODE':str})
# folder_name = "testing"

# premium_receivable_function(df_premium_receivable, folder_name)
# df_premium_receivable.to_excel("export_test.xlsx", index=False)
