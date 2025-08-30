import os
import pandas as pd


def get_insured_name(df_premium, df_claims):
    """
    Build a pivot table mapping (Origin Office, Follower Office Code) -> Name of insured.

    Args:
        df_premium (pd.DataFrame): Premium dataframe (may be empty)
        df_claims (pd.DataFrame): Claims dataframe (may be empty)

    Returns:
        pd.DataFrame: Pivot table with 'Name of insured' aggregated
    """

    # Collect valid frames
    frames = []
    for df in (df_premium, df_claims):
        if not df.empty:
            frames.append(
                df[["Origin Office", "Follower Office Code", "Name of insured"]]
            )

    if not frames:
        # no data at all
        return pd.DataFrame()

    # Merge them if both exist
    df_all = pd.concat(frames, ignore_index=True).drop_duplicates()

    # Pivot with custom aggregator
    df_pivot_insured = df_all.pivot_table(
        index=["Origin Office", "Follower Office Code"],
        values="Name of insured",
        aggfunc=lambda x: ", ".join(x.str.replace(",", "")),
    )

    return df_pivot_insured


def generate_pivot_table(excel_filename, company_name, path_string_wip=None):
    """
    Generate a pivot summary for claims and premium by company.

    Args:
        excel_filename: Path to Excel file containing CR (claims) and PP (premium) sheets
        company_name: Company identifier
        path_string_wip: Optional output folder for saving summary file

    Returns:
        pd.DataFrame: Final pivot summary
    """

    def try_read(sheet, value_col):
        """Helper to read sheet and pivot by Origin & Follower office."""
        try:
            df = pd.read_excel(
                excel_filename,
                sheet_name=sheet,
                converters={"Follower Office Code": str, "Origin Office": str},
            )
            return pd.pivot_table(
                df,
                index=["Origin Office", "Follower Office Code"],
                values=value_col,
                aggfunc="sum",
            ), df
        except ValueError:
            print(f"No {sheet} sheet for {company_name}")
            return pd.DataFrame(), pd.DataFrame()

    # --- Step 1: Read premium & claims sheets ---
    df_pivot_claims, df_claims = try_read("CR", "Total claim amount")
    df_pivot_premium, df_premium = try_read("PP", "Net Premium payable")

    pd.set_option("display.float_format", "{:,.2f}".format)

    # --- Step 2: Merge logic ---
    if not df_pivot_premium.empty and not df_pivot_claims.empty:
        # both present
        df_insured = get_insured_name(df_premium, df_claims)
        df_summary = df_pivot_premium.merge(
            df_pivot_claims, left_index=True, right_index=True, how="outer"
        ).merge(df_insured, left_index=True, right_index=True, how="outer")
        df_summary.reset_index(inplace=True)
        df_summary.fillna(0, inplace=True)

        net_amount = (
            df_summary["Total claim amount"].sum()
            - df_summary["Net Premium payable"].sum()
        )
        if net_amount > 0:
            df_summary["Net Receivable by UIIC"] = (
                df_summary["Total claim amount"] - df_summary["Net Premium payable"]
            )
        else:
            df_summary["Net Payable by UIIC"] = (
                df_summary["Net Premium payable"] - df_summary["Total claim amount"]
            )

        df_summary.rename(
            columns={"Total claim amount": "Claims receivable"}, inplace=True
        )

    elif not df_pivot_premium.empty:  # premium only
        df_insured = get_insured_name(df_premium, pd.DataFrame())
        df_summary = df_pivot_premium.merge(
            df_insured, left_index=True, right_index=True, how="outer"
        )
        df_summary.reset_index(inplace=True)

    elif not df_pivot_claims.empty:  # claims only
        df_insured = get_insured_name(pd.DataFrame(), df_claims)
        df_summary = df_pivot_claims.merge(
            df_insured, left_index=True, right_index=True, how="outer"
        )
        df_summary.reset_index(inplace=True)

    else:  # neither
        print(f"No premium or claims data for {company_name}")
        return pd.DataFrame()

    # --- Step 3: Column reordering (put "Name of insured" last) ---
    if "Name of insured" in df_summary.columns:
        cols = df_summary.columns.tolist()
        idx = cols.index("Name of insured")
        cols = cols[:idx] + cols[idx + 1 :] + cols[idx : idx + 1]
        df_summary = df_summary[cols]

    # --- Step 4: Add totals row ---
    df_summary.loc["Total"] = df_summary.sum(numeric_only=True, axis=0)

    print(df_summary)

    # --- Step 5: Save Excel ---
    if path_string_wip:
        out_file = os.path.join(
            path_string_wip, company_name, f"{company_name}_summary.xlsx"
        )
    else:
        out_file = f"{company_name}_summary.xlsx"

    with pd.ExcelWriter(out_file, engine="xlsxwriter") as writer:
        df_summary.to_excel(writer, sheet_name="Summary", index=False)
        wb, ws = writer.book, writer.sheets["Summary"]

        fmt_currency = wb.add_format({"num_format": "##,##,#0"})
        fmt_bold = wb.add_format({"num_format": "##,##,#", "bold": True})
        fmt_header = wb.add_format({"bold": True, "text_wrap": True, "valign": "top"})

        ws.set_column("E:E", 12, fmt_bold)
        ws.set_column("C:D", 12, fmt_currency)
        ws.set_row(-1, 12, fmt_bold)
        ws.set_row(0, None, fmt_header)

    return df_summary


