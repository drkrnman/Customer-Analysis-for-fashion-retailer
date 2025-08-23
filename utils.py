import pandas as pd
import os
import seaborn as sns
import matplotlib.pyplot as plt
import numpy as np
from matplotlib.ticker import FuncFormatter
from scipy import stats
from docx import Document
import textwrap

# ===== Display options =====
pd.set_option('display.max_colwidth', 15)
pd.set_option('display.width', 100)
pd.set_option('display.expand_frame_repr', True)

# ===== Helpers =====

def read_file(filename, file_type, base_dir=None):
    """Read CSV or DOCX from base_dir (defaults to script dir)."""
    try:
        if base_dir is None:
            base_dir = os.path.dirname(os.path.abspath(__file__))
        file_path = os.path.join(base_dir, filename)
        if file_type == 'csv':
            return pd.read_csv(file_path)
        if file_type == 'docx':
            doc = Document(file_path)
            return [para.text for para in doc.paragraphs]
        print(f"Error: Unsupported file type '{file_type}'.")
        return pd.DataFrame() if file_type == 'csv' else []
    except FileNotFoundError:
        print(f"Error: File '{filename}' not found in {base_dir}.")
        return pd.DataFrame() if file_type == 'csv' else []
    except Exception as e:
        print(f"Error reading '{filename}': {e}")
        return pd.DataFrame() if file_type == 'csv' else []

customers = read_file('customer_stats.csv', 'csv')


def get_user_choice(prompt, options, allow_exit=True):
    """Get validated input against options (dict or list)."""
    if not options:
        print("Error: No options available.")
        return None
    while True:
        print(prompt)
        if isinstance(options, dict):
            for key, value in options.items():
                print(f"{key}. {value}")
        else:
            for i, option in enumerate(options, 1):
                print(f"{i}. {option}")
        if allow_exit:
            print("0. Exit" if isinstance(options, dict) else f"{len(options) + 1}. Exit")
        choice = input("Choose an option: ").strip().lower()
        if allow_exit and (choice == 'exit' or choice == '0' or (isinstance(options, list) and choice == str(len(options) + 1))):
            return None
        try:
            choice = int(choice)
            if isinstance(options, dict) and choice in options:
                return options[choice]
            if isinstance(options, list) and 1 <= choice <= len(options):
                return options[choice - 1]
            print("Invalid option.")
        except ValueError:
            print("Please enter a number or 'exit'.")

# =====================
# Formatting helpers
# =====================

def format_int(x):
    try:
        return f"{int(round(float(x)))}"
    except Exception:
        return str(x)


def format_int_thousands(x):
    try:
        x = float(x)
        return f"{x/1000:.0f} K." if x >= 1000 else f"{int(round(x))}"
    except Exception:
        return str(x)


def format_percent(x):
    try:
        return f"{int(round(float(x)))}% "
    except Exception:
        return f"{str(x)} "


def format_float(x):
    try:
        return f"{float(x):.1f}"
    except Exception:
        return str(x)


def wrap_label(s, width=24):
    s = str(s)
    return textwrap.fill(s, width=width, break_long_words=False)

# =====================
# Visualization functions
# =====================

def create_bar_plot(data, title, formatters, figsize=(14, 6), show=True):
    try:
        data_plot = data.drop("Total", errors="ignore").sort_index(ascending=False)
        if data_plot.empty:
            raise ValueError("No data to plot")

        fig, axes = plt.subplots(1, len(data_plot.columns), figsize=figsize,
                                 sharey=True, constrained_layout=True)
        axes = np.atleast_1d(axes)
        cats = [wrap_label(str(c), 28) for c in data_plot.index]
        colors = plt.cm.tab10.colors

        for ax, col, fmt in zip(axes, data_plot.columns, formatters):
            vals = data_plot[col].astype(float).fillna(0)
            if fmt is format_percent:
                vals = vals * 100
            bars = ax.barh(cats, vals, color=colors[:len(cats)])
            ax.set_title(wrap_label(col, 28), fontsize=12)
            ax.grid(True, linestyle=":", alpha=0.6)
            ax.tick_params(axis="y", labelsize=10)

            # подписи сразу ставим
            for bar, v in zip(bars, vals):
                if not np.isfinite(v):
                    continue
                ha = "left" if v >= 0 else "right"
                pad = 0.01 * (abs(vals.max()) or 1)
                ax.text(v + (pad if v >= 0 else -pad),
                        bar.get_y() + bar.get_height()/2,
                        fmt(v), va="center", ha=ha, fontsize=9)

        fig.suptitle(wrap_label(title, 60), fontsize=14)

        # function that recalculates xlim on each redraw
        def adjust_limits(event):
            for ax in axes:
                rightmost = max([bar.get_width() for bar in ax.patches] + [0])
                renderer = event.renderer
                for t in ax.texts:
                    bb = t.get_window_extent(renderer=renderer)
                    x_data = ax.transData.inverted().transform((bb.x1, 0))[0]
                    rightmost = max(rightmost, x_data)
                x0, _ = ax.get_xlim()
                ax.set_xlim(x0, rightmost * 1.05)

        fig.canvas.mpl_connect("draw_event", adjust_limits)
        if show:
            plt.show()
        return fig

    except Exception as e:
        print(f"Error while plotting the graph: {e}")
        return None


def create_line_plot(metric_ltv, metric_returned_cust, title, index_name, figsize=(16, 9), show=True):
    try:
        fig, (ax1, ax2) = plt.subplots(2, 1, figsize=figsize, sharex=True)

        # LTV lines
        for category in metric_ltv.index:
            ax1.plot(metric_ltv.columns, metric_ltv.loc[category], marker='o', linewidth=2, label=str(category))
        ax1.set_title('LTV (Revenue per 1 customer) in 6 months', pad=10, fontsize=13)
        ax1.grid(True, linestyle='--', alpha=0.6)
        ax1.tick_params(axis='x', rotation=45, labelsize=10)

        # Retention lines
        for category in metric_returned_cust.index:
            ax2.plot(metric_returned_cust.columns, metric_returned_cust.loc[category], marker='o', linewidth=2, label=str(category))
        ax2.set_title('Retention rate (percentage of returned customers) in 6 months', pad=10, fontsize=13)
        ax2.grid(True, linestyle='--', alpha=0.6)
        ax2.tick_params(axis='x', rotation=45, labelsize=10)
        ax2.set_xlabel('Cohort month', fontsize=12)

        # Create a single legend OUTSIDE on the right, wrap long labels
        handles, labels = ax1.get_legend_handles_labels()
        if not handles:
            handles, labels = ax2.get_legend_handles_labels()
        display_labels = [wrap_label(l, 26) for l in labels]
        # Reserve right margin and place legend there
        fig.subplots_adjust(top=0.84, right=0.76, hspace=0.35)
        fig.legend(
            handles,
            display_labels,
            title=wrap_label(index_name, 26),
            loc='center left',
            bbox_to_anchor=(0.78, 0.5),
            fontsize=9,
            frameon=False,
        )

        fig.suptitle(wrap_label(title, 68), fontsize=16)
        if show:
            plt.show()
        return fig
    except Exception as e:
        print(f"Error while plotting line charts: {e}")
        return None


def create_pie_plot(data, title, figsize=(16, 6), show=True):
    """Two pies with a shared figure legend OUTSIDE; nothing gets clipped."""
    try:
        fig, (ax1, ax2) = plt.subplots(1, 2, figsize=figsize)

        fig.subplots_adjust(right=0.74, top=0.84, wspace=0.2)
        fig.suptitle(wrap_label(title, 60), fontsize=16)

        labels_wrapped = [wrap_label(str(lbl), 24) for lbl in data.index]

        ax1.set_title('Percentage of revenue', fontsize=13, pad=8)
        wedges1, texts1, autotexts1 = ax1.pie(
            data['Pers of revenue'],
            labels=None,  # keep labels in legend only (cleaner)
            autopct='%.0f%%',
            startangle=90,
            textprops={'fontsize': 10},
        )

        ax2.set_title('Percentage of customers', fontsize=13, pad=8)
        wedges2, texts2, autotexts2 = ax2.pie(
            data['Pers of customers'],
            labels=None,
            autopct='%.0f%%',
            startangle=90,
            textprops={'fontsize': 10},
        )

        fig.legend(
            wedges1,
            labels_wrapped,
            title=wrap_label(data.index.name or 'Category', 24),
            loc='center left',
            bbox_to_anchor=(0.76, 0.5),
            frameon=False,
            fontsize=10,
        )

        if show:
            plt.show()
        return fig
    except Exception as e:
        print(f"Error while plotting pie charts: {e}")
        return None

# =====================
# Strings / menu
# =====================

columns_str = {
    'age_group': 'Customer Age',
    'gender': 'Customer Gender',
    'first_payment_method': 'First purchase payment method',
    'first_currency': 'First purchase currency',
    'customer_country': 'Customer Country',
    'first_purchase_sum_group': 'First purchases value',
    'first_purchase_prods_cnt_group': 'First purchases items number',
    'store_country': 'Stores country',
}


def select_section():
    columns_dict = {i: key for i, key in enumerate(columns_str.keys(), start=1)}
    choice = get_user_choice("To select a column name enter its number:", columns_dict)
    return (choice, columns_str[choice]) if choice else (None, None)

# =====================
# 1. Executive summary
# =====================

def read_summary():
    text = read_file('Executive_summary.docx', 'docx')
    if text:
        print("\n".join(text))

# =====================
# 2. LTV Factors
# =====================

def ltv_factors(df=customers):
    while True:
        column_name, str_name = select_section()
        if not column_name:
            break
        try:
            agg_funcs = {
                'first_purchase_sum': 'sum',
                'next_sum': 'sum',
                'customer_id': 'count',
                'returned_customer': 'sum',
                'next_purchases_cnt': 'sum',
            }
            metrics = df.pivot_table(index=column_name, aggfunc=agg_funcs)
            metrics.loc['Total'] = metrics.sum()

            metrics['LTV'] = ((metrics['first_purchase_sum'] + metrics['next_sum']) / metrics['customer_id']).round(2)
            metrics['Num of cust'] = metrics['customer_id'].round(0)
            metrics['Pers of cust'] = (metrics['customer_id'] / len(df) * 100).round(1)
            metrics['Perc rep cust'] = (metrics['returned_customer'] / metrics['customer_id'] * 100).round(1)
            metrics['Avg num pur'] = (metrics['next_purchases_cnt'] / metrics['returned_customer']).round(1)
            metrics['First pur'] = (metrics['first_purchase_sum'] / metrics['customer_id']).round(2)
            metrics['Rep pur'] = (metrics['next_sum'] / metrics['returned_customer']).round(2)

            metrics = metrics[
                ['LTV', 'Num of cust', 'Pers of cust', 'Perc rep cust', 'Avg num pur', 'First pur', 'Rep pur']
            ]
            print(f'LTV factors. Split by {str_name}.')
            print(metrics)

            create_bar_plot(
                metrics,
                f'LTV factors. Split by {str_name}.',
                [format_float, format_int_thousands, format_percent, format_percent, format_float, format_int, format_int],
                figsize=(28, 12),
            )
        except Exception as e:
            print(f"Error: {e}")

# =====================
# 3. LTV Cohorts
# =====================

def ltv_cohort(df=customers):
    while True:
        column_name, str_name = select_section()
        if not column_name:
            break
        try:
            agg_funcs = {
                'first_purchase_sum': 'sum',
                'customer_id': 'count',
                'returned_customer': 'sum',
                'next_sum': 'sum',
            }
            metrics = df.pivot_table(
                index=column_name,
                columns='cohort_month',
                values=list(agg_funcs.keys()),
                aggfunc=agg_funcs,
            )
            metric_ltv = ((metrics['first_purchase_sum'] + metrics['next_sum']) / metrics['customer_id']).round(2)
            metric_returned_cust = (metrics['returned_customer'] / metrics['customer_id']).round(2)

            print(f'LTV dynamics split by {str_name}.')
            print(metric_ltv)
            print(metric_returned_cust)

            create_line_plot(
                metric_ltv,
                metric_returned_cust,
                f'LTV dynamics split by {str_name}.',
                column_name,
                figsize=(16, 9),
            )
        except Exception as e:
            print(f"Error: {e}")

# =====================
# 4. Revenue Structure
# =====================

def revenue_structure(df=customers):
    while True:
        column_name, str_name = select_section()
        if not column_name:
            break
        try:
            agg_funcs = {'first_purchase_sum': 'sum', 'next_sum': 'sum', 'customer_id': 'count'}
            metrics = df.pivot_table(index=column_name, aggfunc=agg_funcs)
            total_revenue = df['first_purchase_sum'].sum() + df['next_sum'].sum()
            total_customers = len(df)

            metrics['Pers of revenue'] = ((metrics['first_purchase_sum'] + metrics['next_sum']) / total_revenue * 100).round(1)
            metrics['Pers of customers'] = (metrics['customer_id'] / total_customers * 100).round(1)
            metrics = metrics[['Pers of revenue', 'Pers of customers']]

            print(f'Revenue structure split by {str_name}.')
            print(metrics)

            create_pie_plot(metrics, f'Distribution by {str_name}')
        except Exception as e:
            print(f"Error: {e}")

# =====================
# 5. Statistical Tests
# =====================

def statistical_tests():
    options = {1: 'Chi-square test', 2: 'T-test'}
    while True:
        choice = get_user_choice("\nStatistical tests:", options)
        if choice is None:
            break
        if choice == 'Chi-square test':
            chi2_menu()
        elif choice == 'T-test':
            ttest_menu()


def chi2_menu():
    options = {1: 'By countries', 2: 'By payment methods'}
    while True:
        choice = get_user_choice("\nChi-square test options:", options)
        if choice is None:
            break
        if choice == 'By countries':
            chi2_custom('returned_customer', 'Returned customer', 'customer_country', 'Customer Country', customers)
        elif choice == 'By payment methods':
            chi2_custom('returned_customer', 'Returned customer', 'first_payment_method', 'First purchase payment method', customers)


def ttest_menu():
    while True:
        countries = list(customers['customer_country'].unique())
        if not countries:
            print("Error: No countries available in the dataset.")
            break
        choice1 = get_user_choice("\nAvailable countries:", countries)
        if choice1 is None:
            break
        countries2 = [c for c in countries if c != choice1]
        if not countries2:
            print("Error: No other countries available for comparison.")
            break
        choice2 = get_user_choice("\nSelect second country:", countries2)
        if choice2 is None:
            break
        if choice1 == choice2:
            print("Error: You cannot select the same country twice.")
            continue
        try:
            t_test_custom(customers, 'returned_customer', 'Returned customer', 'customer_country', 'Customer Country', choice1, choice2)
        except Exception as e:
            print(f"Error running T-test: {e}")

# =====================
# Test functions
# =====================

def chi2_custom(groups, groups_name, columns, columns_name, df):
    try:
        contingency_table = pd.crosstab(df[groups], df[columns])
        contingency_table_percent = pd.crosstab(df[groups], df[columns], normalize='columns') * 100
        contingency_table_percent = contingency_table_percent.round(0)

        chi2_stat, p_value, dof, expected = stats.chi2_contingency(contingency_table)

        print(f"Does percentage of {groups_name} differ across {columns_name}?")
        print('Number of customers:', '\n')
        print(contingency_table)
        print(f"\nNumber of customers. % of totals by {columns_name}:")
        print(contingency_table_percent, '\n')

        print(f"Null hypothesis: {groups_name} distribution is independent of  {columns_name}.")
        print('P-value of Chi-square test = ', p_value.round(3))
        if p_value < 0.05:
            print('We reject Null hypothesis.', '\n')
            print(f"There is statistical evidence that {groups_name} distribution differs across {columns_name}.")
            
        else:
            print('We fail to reject the null hypothesis', '\n')
            print(f"There is NO statistical evidence that {groups_name} distribution differs across {columns_name}.")
           
    except Exception as e:
        print(f"Error in chi2_custom: {e}")


def t_test_custom(df, groups, groups_name, columns, columns_name, group_1, group_2):
    try:
        contingency_table = pd.crosstab(df[groups], df[columns])
        # Percent of True per column
        if True in contingency_table.index:
            percent_true = (contingency_table.loc[True] / contingency_table.sum()) * 100
        elif 1 in contingency_table.index:
            percent_true = (contingency_table.loc[1] / contingency_table.sum()) * 100
        else:
            # Fallback: treat any non-zero as True
            true_row = contingency_table.loc[(contingency_table.index != 0)].sum()
            percent_true = (true_row / contingency_table.sum()) * 100

        customers_group1 = df[df[columns] == group_1][groups]
        customers_group2 = df[df[columns] == group_2][groups]

        t_stat, p_value = stats.ttest_ind(customers_group1, customers_group2)
        print(f"Is there a significant difference between percentage of {groups_name} for {group_1} and {group_2}?\n")
        print('Number of customers:', '\n')
        print(contingency_table, '\n')

        print(f"Percentage of {groups_name}:")
        print(f"for {group_1}  = {str(percent_true[group_1].round(2))}")
        print(f"for {group_1} = {str(percent_true[group_2].round(2))}\n")

        print(f"Null hypothesis: percentage of returned customers is the same for {group_1} and {group_2}.")
        print('P-value of t-test (for independent samples) = ', p_value.round(3))
        if p_value < 0.05:
            print('We reject Null hypothesis.', '\n')
            print(f"There is statistical evidence that percentage of {groups_name} differs for {group_1} and {group_2}.")
        else:
            print('We fail to reject the null hypothesis', '\n')
            print(f"There is NO statistically significant evidence that {groups_name}differs for {group_1} and {group_2}.")
    except Exception as e:
        print(f"Error in t_test_custom: {e}")

# =====================
# GUI-friendly computation helpers (no prints, no input)
# =====================

def compute_ltv_factors_for_column(df, column_name):
    """Compute LTV factors table for a specific column; returns (metrics_df, title, formatters)."""
    agg_funcs = {
        'first_purchase_sum': 'sum',
        'next_sum': 'sum',
        'customer_id': 'count',
        'returned_customer': 'sum',
        'next_purchases_cnt': 'sum',
    }
    metrics = df.pivot_table(index=column_name, aggfunc=agg_funcs)
    metrics.loc['Total'] = metrics.sum()

    metrics['LTV'] = ((metrics['first_purchase_sum'] + metrics['next_sum']) / metrics['customer_id']).round(2)
    metrics['Num of cust'] = metrics['customer_id'].round(0)
    metrics['Pers of cust'] = (metrics['customer_id'] / len(df) * 100).round(1)
    metrics['Perc rep cust'] = (metrics['returned_customer'] / metrics['customer_id'] * 100).round(1)
    metrics['Avg num pur'] = (metrics['next_purchases_cnt'] / metrics['returned_customer']).round(1)
    metrics['First pur'] = (metrics['first_purchase_sum'] / metrics['customer_id']).round(2)
    metrics['Rep pur'] = (metrics['next_sum'] / metrics['returned_customer']).round(2)

    metrics = metrics[
        ['LTV', 'Num of cust', 'Pers of cust', 'Perc rep cust', 'Avg num pur', 'First pur', 'Rep pur']
    ]
    title = f"LTV factors. Split by {columns_str.get(column_name, column_name)}."
    formatters = [format_float, format_int_thousands, format_percent, format_percent, format_float, format_int, format_int]
    return metrics, title, formatters


def compute_ltv_cohort_for_column(df, column_name):
    """Compute LTV cohort dynamics; returns (metric_ltv_df, metric_returned_cust_df, title, index_name)."""
    agg_funcs = {
        'first_purchase_sum': 'sum',
        'customer_id': 'count',
        'returned_customer': 'sum',
        'next_sum': 'sum',
    }
    metrics = df.pivot_table(
        index=column_name,
        columns='cohort_month',
        values=list(agg_funcs.keys()),
        aggfunc=agg_funcs,
    )
    metric_ltv = ((metrics['first_purchase_sum'] + metrics['next_sum']) / metrics['customer_id']).round(2)
    metric_returned_cust = (metrics['returned_customer'] / metrics['customer_id']).round(2)

    title = f"LTV dynamics split by {columns_str.get(column_name, column_name)}."
    index_name = column_name
    return metric_ltv, metric_returned_cust, title, index_name


def compute_revenue_structure_for_column(df, column_name):
    """Compute revenue structure metrics; returns (metrics_df, title)."""
    agg_funcs = {'first_purchase_sum': 'sum', 'next_sum': 'sum', 'customer_id': 'count'}
    metrics = df.pivot_table(index=column_name, aggfunc=agg_funcs)
    total_revenue = df['first_purchase_sum'].sum() + df['next_sum'].sum()
    total_customers = len(df)

    metrics['Pers of revenue'] = ((metrics['first_purchase_sum'] + metrics['next_sum']) / total_revenue * 100).round(1)
    metrics['Pers of customers'] = (metrics['customer_id'] / total_customers * 100).round(1)
    metrics = metrics[['Pers of revenue', 'Pers of customers']]

    title = f"Distribution by {columns_str.get(column_name, column_name)}"
    return metrics, title


def compute_chi2_result(df, groups, groups_name, columns, columns_name):
    """Compute chi-square test artifacts and textual interpretation."""
    contingency_table = pd.crosstab(df[groups], df[columns])
    contingency_table_percent = (pd.crosstab(df[groups], df[columns], normalize='columns') * 100).round(0)
    chi2_stat, p_value, dof, expected = stats.chi2_contingency(contingency_table)

    null_hypothesis = f"Null hypothesis: {groups_name} distribution is independent of {columns_name}."
    decision = 'reject' if p_value < 0.05 else 'fail_to_reject'
    interpretation = (
        f"There is statistical evidence that {groups_name} distribution differs across {columns_name}."
        if decision == 'reject'
        else f"There is NO statistical evidence that {groups_name} distribution differs across {columns_name}."
    )
    return {
        'contingency_table': contingency_table,
        'contingency_table_percent': contingency_table_percent,
        'p_value': float(np.round(p_value, 3)),
        'decision': decision,
        'null_hypothesis': null_hypothesis,
        'interpretation': interpretation,
    }


def compute_ttest_result(df, groups, groups_name, columns, columns_name, group_1, group_2):
    """Compute independent t-test artifacts and textual interpretation for two groups."""
    contingency_table = pd.crosstab(df[groups], df[columns])
    if True in contingency_table.index:
        percent_true = (contingency_table.loc[True] / contingency_table.sum()) * 100
    elif 1 in contingency_table.index:
        percent_true = (contingency_table.loc[1] / contingency_table.sum()) * 100
    else:
        true_row = contingency_table.loc[(contingency_table.index != 0)].sum()
        percent_true = (true_row / contingency_table.sum()) * 100

    customers_group1 = df[df[columns] == group_1][groups]
    customers_group2 = df[df[columns] == group_2][groups]

    t_stat, p_value = stats.ttest_ind(customers_group1, customers_group2)

    null_hypothesis = f"Null hypothesis: percentage of returned customers is the same for {group_1} and {group_2}."
    decision = 'reject' if p_value < 0.05 else 'fail_to_reject'
    interpretation = (
        f"There is statistical evidence that percentage of {groups_name} differs for {group_1} and {group_2}."
        if decision == 'reject'
        else f"There is NO statistically significant evidence that {groups_name} differs for {group_1} and {group_2}."
    )

    return {
        'contingency_table': contingency_table,
        'percent_true': percent_true,
        'p_value': float(np.round(p_value, 3)),
        'decision': decision,
        'null_hypothesis': null_hypothesis,
        'interpretation': interpretation,
        'group_1': group_1,
        'group_2': group_2,
        'groups_name': groups_name,
        'columns_name': columns_name,
    }
