import streamlit as st
import pandas as pd
import os
from main import (
    load_data,
    clean_data,
    analyze_data,
    export_report,
    generate_charts,
    embed_charts_in_excel,
)


def main():
    st.set_page_config(page_title="Household Expense Tracker", layout="wide")

    st.title("🏠 Household Expense Tracker")
    st.markdown("---")

    # Sidebar navigation
    st.sidebar.title("Navigation")
    page = st.sidebar.selectbox("Choose a page", ["View Dataset", "Generate Analysis Report", "View Charts"])

    # Load data once
    if "expense_data" not in st.session_state:
        with st.spinner("Loading data..."):
            try:
                df = load_data("household_expenses.xlsx")
                df_clean = clean_data(df)
                st.session_state.expense_data = df_clean
                st.success("Data loaded successfully!")
            except Exception as e:
                st.error(f"Error loading data: {str(e)}")
                return

    df = st.session_state.expense_data

    # Page 1: View Dataset
    if page == "View Dataset":
        st.header("📋 Raw Dataset Preview")
        st.dataframe(df.head(20))
        st.info(f"Total records: {len(df)}")

    # Page 2: Generate Analysis Report
    elif page == "Generate Analysis Report":
        st.header("📊 Generate Analysis Report")

        if st.button("Generate Report"):
            with st.spinner("Generating analysis report..."):
                try:
                    # Perform analysis
                    results = analyze_data(df)

                    # Export report
                    output_excel = "cleaned_expenses.xlsx"
                    export_report(output_excel, results)

                    # Generate charts
                    charts_dir = "charts"
                    chart_paths = generate_charts(df, charts_dir)
                    embed_charts_in_excel(output_excel, chart_paths)

                    st.success("Analysis report generated successfully!")

                    # Display summary
                    st.subheader("Analysis Summary")
                    col1, col2, col3 = st.columns(3)

                    with col1:
                        st.metric("Total Records", len(df))
                        st.metric("Average Expense", f"₹{results['average_expense']:.0f}")

                    with col2:
                        st.metric("Max Expense", f"₹{results['max_expense']}")
                        st.metric("Min Expense", f"₹{results['min_expense']}")

                    with col3:
                        top_cat = results["category_totals"].idxmax()
                        st.metric("Top Category", top_cat)
                        st.metric("Payments by UPI", results["payment_counts"].get("UPI", 0))

                    # Download link
                    if os.path.exists(output_excel):
                        with open(output_excel, "rb") as f:
                            st.download_button(
                                label="📥 Download Expense Report (Excel)",
                                data=f.read(),
                                file_name=output_excel,
                                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                            )

                except Exception as e:
                    st.error(f"Error generating report: {str(e)}")

    # Page 3: View Charts
    elif page == "View Charts":
        st.header("📈 Data Visualization Charts")

        charts_dir = "charts"
        chart_files = {
            "Expenses by Category (Bar Chart)": "category_expenses_bar.png",
            "Payment Mode Distribution (Pie Chart)": "payment_mode_pie.png",
            "Monthly Expenses (Line Chart)": "monthly_expenses_line.png",
            "Expense Distribution (Histogram)": "expense_hist.png",
        }

        if st.button("Generate Charts"):
            with st.spinner("Generating charts..."):
                try:
                    chart_paths = generate_charts(df, charts_dir)
                    st.success("Charts generated successfully!")
                except Exception as e:
                    st.error(f"Error generating charts: {str(e)}")

        # Display charts
        for chart_name, chart_file in chart_files.items():
            chart_path = os.path.join(charts_dir, chart_file)
            if os.path.exists(chart_path):
                st.subheader(chart_name)
                st.image(chart_path, use_column_width=True)
            else:
                st.info(f"{chart_name} not found. Click 'Generate Charts' to create it.")


if __name__ == "__main__":
    main()
