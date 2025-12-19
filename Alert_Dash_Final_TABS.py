import pandas as pd
import tkinter as tk
from tkinter import filedialog, messagebox
import re

# =========================================
# Main processing function
# =========================================
def process_excel():
    file_path = filedialog.askopenfilename(
        title="Select Excel file",
        filetypes=[("Excel files", "*.xlsx *.xls")]
    )
    if not file_path:
        return

    try:
        df = pd.read_excel(file_path)

        # Required columns
        required_columns = ['ipAddress', 'title', 'category', 'applicationName']
        missing_cols = [col for col in required_columns if col not in df.columns]
        if missing_cols:
            messagebox.showerror("Error", f"Missing required columns: {missing_cols}")
            return

        # ---------------------------
        # Clean & normalize title
        # ---------------------------
        df['title_clean'] = (
            df['title']
            .astype(str)
            .str.lower()
            .apply(lambda x: re.sub(r'\d+', '<num>', x))
            .str.replace(r'[^a-z0-9 <>]', ' ', regex=True)
            .str.replace(r'\s+', ' ', regex=True)
            .str.strip()
        )

        # ---------------------------
        # Group & count repeated issues
        # ---------------------------
        issue_counts = (
            df.groupby(['ipAddress', 'applicationName', 'category', 'title_clean'])
              .agg(
                  repeat_count=('title_clean', 'size'),
                  example_title=('title', 'first')
              )
              .reset_index()
        )

        issue_counts = issue_counts[
            ['ipAddress', 'applicationName', 'category',
             'example_title', 'repeat_count', 'title_clean']
        ]

        # Sort globally
        all_issues_sorted = issue_counts.sort_values(
            'repeat_count', ascending=False
        )

        # ---------------------------
        # Category Summary (overall)
        # ---------------------------
        category_summary = (
            issue_counts.groupby('category')
                        .agg(
                            total_issues=('repeat_count', 'sum'),
                            unique_titles=('title_clean', 'nunique')
                        )
                        .reset_index()
                        .sort_values('total_issues', ascending=False)
        )

        # ---------------------------
        # Save output
        # ---------------------------
        output_path = filedialog.asksaveasfilename(
            defaultextension=".xlsx",
            filetypes=[("Excel files", "*.xlsx *.xls")],
            title="Save Issues Report"
        )
        if not output_path:
            return

        with pd.ExcelWriter(output_path, engine='xlsxwriter') as writer:
            workbook = writer.book

            # ---------------------------
            # All Issues Sheet
            # ---------------------------
            all_issues_sorted.to_excel(
                writer, index=False, sheet_name='All_Issues'
            )

            # ---------------------------
            # Category Summary Sheet
            # ---------------------------
            category_summary.to_excel(
                writer, index=False, sheet_name='Category_Summary'
            )

            summary_sheet = writer.sheets['Category_Summary']

            # Chart: Total Issues per Category
            chart1 = workbook.add_chart({'type': 'column'})
            chart1.add_series({
                'name': 'Total Issues',
                'categories': ['Category_Summary', 1, 0,
                               len(category_summary), 0],
                'values':     ['Category_Summary', 1, 1,
                               len(category_summary), 1],
                'data_labels': {'value': True}
            })
            chart1.set_title({'name': 'Total Issues per Category'})
            chart1.set_x_axis({'name': 'Category'})
            chart1.set_y_axis({'name': 'Total Issues'})
            summary_sheet.insert_chart('E2', chart1)

            # ---------------------------
            # Per-Application Sheets + Charts
            # ---------------------------
            for app_name, app_df in all_issues_sorted.groupby('applicationName'):
                sheet_name = f"App_{app_name}"[:31]  # Excel sheet name limit

                app_df.to_excel(
                    writer, index=False, sheet_name=sheet_name
                )

                worksheet = writer.sheets[sheet_name]

                top10 = app_df.head(10)
                top_n = len(top10)

                if top_n == 0:
                    continue

                chart = workbook.add_chart({'type': 'bar'})
                chart.add_series({
                    'name': 'Repeat Count',
                    'categories': [sheet_name, 1, 3, top_n, 3],
                    'values':     [sheet_name, 1, 4, top_n, 4],
                    'data_labels': {'value': True}
                })

                chart.set_title({
                    'name': f'Top 10 Repeated Issues - {app_name}'
                })
                chart.set_x_axis({'name': 'Repeat Count'})
                chart.set_y_axis({'name': 'Issue Title'})

                worksheet.insert_chart('H2', chart)

        messagebox.showinfo(
            "Success",
            f"Issues report created successfully:\n{output_path}"
        )

    except Exception as e:
        messagebox.showerror("Error", str(e))


# =========================================
# GUI Layout
# =========================================
root = tk.Tk()
root.title("Repeated Issues Report with Application Breakdown")
root.geometry("600x260")

label = tk.Label(
    root,
    text=(
        "Select Excel file to generate repeated issues report\n"
        "• Application-wise sheets\n"
        "• Top-10 issues per application\n"
        "• Category summary with charts"
    ),
    font=("Arial", 11)
)
label.pack(pady=25)

btn = tk.Button(
    root,
    text="Select Excel and Process",
    font=("Arial", 12),
    command=process_excel
)
btn.pack(pady=20)

root.mainloop()
