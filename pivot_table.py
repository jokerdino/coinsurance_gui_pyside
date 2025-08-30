import pandas as pd


def get_insured_name(df_premium, df_claims):
    if df_premium.empty:
        df_claims = df_claims.drop_duplicates(
            subset=["Origin Office", "Follower Office Code", "Name of insured"]
        )
        df_pivot_insured = df_claims.pivot_table(
            index=["Origin Office", "Follower Office Code"],
            values="Name of insured",
            aggfunc=lambda x: ", ".join(x.str.replace(",", "")),
        )

    elif df_claims.empty:
        df_premium = df_premium.drop_duplicates(
            subset=["Origin Office", "Follower Office Code", "Name of insured"]
        )
        print(df_premium)
        df_pivot_insured = df_premium.pivot_table(
            index=["Origin Office", "Follower Office Code"],
            values="Name of insured",
            aggfunc=lambda x: ", ".join(x.str.replace(",", "")),
        )

    else:
        df_claims = df_claims.drop_duplicates(
            subset=["Origin Office", "Follower Office Code", "Name of insured"]
        )
        df_premium = df_premium.drop_duplicates(
            subset=["Origin Office", "Follower Office Code", "Name of insured"]
        )
        df_claims = df_claims[
            ["Origin Office", "Follower Office Code", "Name of insured"]
        ]
        df_premium = df_premium[
            ["Origin Office", "Follower Office Code", "Name of insured"]
        ]
        df_concat = pd.concat([df_claims, df_premium])
        df_concat = df_concat.drop_duplicates(
            subset=["Origin Office", "Follower Office Code", "Name of insured"]
        )
        df_pivot_insured = df_concat.pivot_table(
            index=["Origin Office", "Follower Office Code"],
            values="Name of insured",
            aggfunc=lambda x: ", ".join(x.str.replace(",", "")),
        )

    return df_pivot_insured


def generate_pivot_table(excel_filename, company_name, path_string_wip):
    print(f"{company_name=}")
    try:
        df_claims = pd.read_excel(
            excel_filename,
            "CR",
            converters={"Follower Office Code": str, "Origin Office": str},
        )
        df_pivot_claims = pd.pivot_table(
            df_claims,
            index=["Origin Office", "Follower Office Code"],
            values="Total claim amount",
            aggfunc="sum",
        )
    except ValueError:
        df_pivot_claims = pd.DataFrame()
        print("No claims sheet for this company")

    try:
        df_premium = pd.read_excel(
            excel_filename,
            "PP",
            converters={"Follower Office Code": str, "Origin Office": str},
        )
        df_pivot_premium = pd.pivot_table(
            df_premium,
            index=["Origin Office", "Follower Office Code"],
            values="Net Premium payable",
            aggfunc="sum",
        )
    except ValueError:
        df_pivot_premium = pd.DataFrame()
        print("No premium sheet for this company")

    pd.set_option("display.float_format", "{:,.2f}".format)
    if not df_pivot_premium.empty:
        if not df_pivot_claims.empty:
            df_insured_name = get_insured_name(
                df_premium=df_premium, df_claims=df_claims
            )
            df_pivot_merged = pd.merge(
                pd.merge(
                    df_pivot_premium,
                    df_pivot_claims,
                    left_index=True,
                    right_index=True,
                    how="outer",
                ),
                df_insured_name,
                left_index=True,
                right_index=True,
                how="outer",
            )
            df_pivot_merged.reset_index(inplace=True)
            df_pivot_merged.fillna(0, inplace=True)
            net_amount = (
                df_pivot_merged["Total claim amount"].sum()
                - df_pivot_merged["Net Premium payable"].sum()
            )
            if net_amount > 0:
                df_pivot_merged["Net Receivable by UIIC"] = (
                    df_pivot_merged["Total claim amount"]
                    - df_pivot_merged["Net Premium payable"]
                )
                df_pivot_merged.style.format({"Net Receivable by UIIC": "{0:,.2f}"})
            else:
                df_pivot_merged["Net Payable by UIIC"] = (
                    df_pivot_merged["Net Premium payable"]
                    - df_pivot_merged["Total claim amount"]
                )
                df_pivot_merged.style.format({"Net Payable by UIIC": "{0:,.2f}"})
            df_pivot_merged.rename(
                {"Total claim amount": "Claims receivable"}, axis=1, inplace=True
            )
            temp_cols = df_pivot_merged.columns.tolist()
            index = df_pivot_merged.columns.get_loc("Name of insured")
            new_cols = (
                temp_cols[0:index]
                + temp_cols[index + 1 :]
                + temp_cols[index : index + 1]
            )
            df_pivot_merged = df_pivot_merged[new_cols]
        else:
            df_insured_name = get_insured_name(
                df_premium=df_premium, df_claims=pd.DataFrame()
            )
            df_pivot_merged = df_pivot_premium
            df_pivot_merged = df_pivot_merged.merge(
                df_insured_name, left_index=True, right_index=True, how="outer"
            )

            df_pivot_merged.reset_index(inplace=True)

            temp_cols = df_pivot_merged.columns.tolist()
            index = df_pivot_merged.columns.get_loc("Name of insured")
            new_cols = (
                temp_cols[0:index]
                + temp_cols[index + 1 :]
                + temp_cols[index : index + 1]
            )
    else:
        df_insured_name = get_insured_name(
            df_premium=pd.DataFrame(), df_claims=df_claims
        )
        df_pivot_merged = df_pivot_claims
        df_pivot_merged = df_pivot_merged.merge(
            df_insured_name, left_index=True, right_index=True, how="outer"
        )
        df_pivot_merged.reset_index(inplace=True)

        temp_cols = df_pivot_merged.columns.tolist()
        index = df_pivot_merged.columns.get_loc("Name of insured")
        new_cols = (
            temp_cols[0:index] + temp_cols[index + 1 :] + temp_cols[index : index + 1]
        )

    df_pivot_merged.loc["Total"] = df_pivot_merged.sum(numeric_only=True, axis=0)

    print(df_pivot_merged)
    if path_string_wip:
        file_path = f"{path_string_wip}/{company_name}/{company_name}_summary.xlsx"
    else:
        file_path = f"{company_name}_summary.xlsx"
    with pd.ExcelWriter(file_path, engine="xlsxwriter") as writer:
        df_pivot_merged.to_excel(writer, sheet_name="Summary", index=False)
        format_workbook = writer.book

        format_currency = format_workbook.add_format({"num_format": "##,##,#0"})
        format_bold = format_workbook.add_format(
            {"num_format": "##,##,#", "bold": True}
        )
        format_header = format_workbook.add_format(
            {"bold": True, "text_wrap": True, "valign": "top"}
        )
        format_worksheet = writer.sheets["Summary"]

        format_worksheet.set_column("E:E", 12, format_bold)
        format_worksheet.set_column("C:D", 12, format_currency)
        format_worksheet.set_row(-1, 12, format_bold)
        format_worksheet.set_row(0, None, format_header)
        # format_worksheet.autofit()
