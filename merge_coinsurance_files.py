import pandas as pd
from company_wise_reports import worksheet_formatter


def merge_multiple_files(list_file_names, new_file_name):
    output_premium = pd.DataFrame()
    output_claims = pd.DataFrame()

    bool_claims, bool_premium = False, False
    for file in list_file_names:
        try:
            df_premium = pd.read_excel(
                file, sheet_name="PP", converters={"Follower Office Code": str}
            )
            output_premium = pd.concat([output_premium, df_premium])
            bool_premium = True

        except ValueError as e:
            pass
        try:
            df_claims = pd.read_excel(
                file, sheet_name="CR", converters={"Follower Office Code": str}
            )
            output_claims = pd.concat([output_claims, df_claims])
            bool_claims = True
        except ValueError as e:
            pass

    if bool_claims and bool_premium:
        with pd.ExcelWriter(
            f"{new_file_name}.xlsx", datetime_format="dd/mm/yyyy"
        ) as writer:
            output_claims.sort_values("Voucher date").to_excel(
                writer, index=False, sheet_name="CR"
            )
            output_premium.sort_values("Accounting date").to_excel(
                writer, index=False, sheet_name="PP"
            )

            worksheet_formatter(writer, "PP")
            worksheet_formatter(writer, "CR")

    elif bool_claims:
        with pd.ExcelWriter(
            f"{new_file_name}.xlsx", datetime_format="dd/mm/yyyy"
        ) as writer:
            output_claims.sort_values("Voucher date").to_excel(
                writer, index=False, sheet_name="CR"
            )
            worksheet_formatter(writer, "CR")

    elif bool_premium:
        with pd.ExcelWriter(
            f"{new_file_name}.xlsx", datetime_format="dd/mm/yyyy"
        ) as writer:
            output_premium.sort_values("Accounting date").to_excel(
                writer, index=False, sheet_name="PP"
            )
            worksheet_formatter(writer, "PP")
    else:
        pass
