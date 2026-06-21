# DE-PY
convert file DE PY 
import pdfplumber
import pandas as pd
from pathlib import Path

input_folder = r"C:\Users\Ritesh\Downloads\Payment Stub_NY\DE_AG\DE PY\Data"
output_file = r"C:\Users\Ritesh\Downloads\Payment Stub_NY\DE_AG\DE PY\Final_Output.xlsx"

all_output = []

for pdf_file in Path(input_folder).glob("*.pdf"):

    print(f"Processing: {pdf_file.name}")

    payment_rows = []
    calc_rows = []

    with pdfplumber.open(pdf_file) as pdf:

        current_section = ""

        for page in pdf.pages:

            tables = page.extract_tables()

            if not tables:
                continue

            for table in tables:

                i = 0

                while i < len(table):

                    row = table[i]

                    if not row:
                        i += 1
                        continue

                    row_text = " ".join(
                        [str(x).strip() for x in row if x]
                    )

                    # Detect Section
                    if "Payment Calculation Details" in row_text:
                        current_section = "CALC"
                        i += 1
                        continue

                    if "Payments" in row_text:
                        current_section = "PAYMENT"
                        i += 1
                        continue

                    # =========================
                    # PAYMENT TABLE
                    # =========================
                    if current_section == "PAYMENT":

                        try:

                            if (
                                len(row) >= 9
                                and row[1]
                                and str(row[1]).strip().isdigit()
                            ):

                                payment_rows.append({
                                    "Child Name": str(row[0]).strip(),
                                    "Client ID": str(row[1]).strip(),
                                    "DOB": row[2],
                                    "Att Month": row[3],
                                    "Days Billed": row[4],
                                    "Paid Days Under 2": row[5],
                                    "Paid Day Over 2": row[6],
                                    "Parent Fee/Copay": row[7],
                                    "Amount Paid": row[8]
                                })

                        except:
                            pass

                    # =========================
                    # CALC TABLE
                    # =========================
                    elif current_section == "CALC":

                        try:

                            if (
                                len(row) >= 11
                                and row[1]
                                and str(row[1]).strip().isdigit()
                            ):

                                formula = ""

                                # next row formula
                                if i + 1 < len(table):

                                    next_row = table[i + 1]

                                    formula_parts = []

                                    for cell in next_row:

                                        if cell:
                                            formula_parts.append(
                                                str(cell).replace("\n", " ")
                                            )

                                    formula = " ".join(formula_parts)

                                calc_rows.append({
                                    "Client ID": str(row[1]).strip(),
                                    "Half Day Rate": row[2],
                                    "Full Day Rate": row[3],
                                    "Half Days": row[4],
                                    "Full Days": row[5],
                                    "1 and Half Days": row[6],
                                    "Absent Days": row[7],
                                    "Holidays": row[8],
                                    "Auth Type": row[9],
                                    "Extended Care": row[10],
                                    "Formula": formula
                                })

                        except:
                            pass

                    i += 1

    # =========================
    # Merge
    # =========================

    df_payment = pd.DataFrame(payment_rows)
    df_calc = pd.DataFrame(calc_rows)

    if len(df_payment) == 0:
        continue

    if len(df_calc) > 0:

        df_payment["Client ID"] = (
            df_payment["Client ID"]
            .astype(str)
            .str.strip()
        )

        df_calc["Client ID"] = (
            df_calc["Client ID"]
            .astype(str)
            .str.strip()
        )

        final_df = pd.merge(
            df_payment,
            df_calc,
            on="Client ID",
            how="left"
        )

    else:
        final_df = df_payment

    final_df["Source PDF"] = pdf_file.name

    all_output.append(final_df)

# =========================
# Save Output
# =========================

if all_output:

    final_output = pd.concat(
        all_output,
        ignore_index=True
    )

    final_output.to_excel(
        output_file,
        index=False
    )

    print("\nDone")
    print(output_file)

else:

    print("No Data Found")
