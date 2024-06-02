import pandas as pd
import openpyxl
import numpy as np
import matplotlib.pyplot as plt # type: ignore
import statsmodels.api as sm

def data_frame_from_xlsx(xlsx_file, range_name):
    """ Get a single rectangular region from the specified file.
    range_name can be a standard Excel reference ('Sheet1!A2:B7') or 
    refer to a named region ('my_cells')."""
    wb = openpyxl.load_workbook(xlsx_file, data_only=True, read_only=True)
    if '!' in range_name:
        # passed a worksheet!cell reference
        ws_name, reg = range_name.split('!')
        if ws_name.startswith("'") and ws_name.endswith("'"):
            # optionally strip single quotes around sheet name
            ws_name = ws_name[1:-1]
        region = wb[ws_name][reg]
    else:
        # passed a named range; find the cells in the workbook
        full_range = wb.get_named_range(range_name)
        if full_range is None:
            raise ValueError(
                'Range "{}" not found in workbook "{}".'.format(range_name, xlsx_file)
            )
        # convert to list (openpyxl 2.3 returns a list but 2.4+ returns a generator)
        destinations = list(full_range.destinations) 
        if len(destinations) > 1:
            raise ValueError(
                'Range "{}" in workbook "{}" contains more than one region.'
                .format(range_name, xlsx_file)
            )
        ws, reg = destinations[0]
        # convert to worksheet object (openpyxl 2.3 returns a worksheet object 
        # but 2.4+ returns the name of a worksheet)
        if isinstance(ws, str):
            ws = wb[ws]
        region = ws[reg]
    # an anonymous user suggested this to catch a single-cell range (untested):
    # if not isinstance(region, 'tuple'): df = pd.DataFrame(region.value)
    df = pd.DataFrame([cell.value for cell in row] for row in region)
    return df

def first_row_as_columns(df):
    '''This function turns the first row in the dataframe to its columns'''
    df.columns = df.iloc[0]
    df = df.iloc[2:, :]

    return df

def transform_dates(df, drop_list = [], needed_months = [], drop_last_data = 'Y'):
    '''Since all the dates are different in different datasets, this function transforms
      the dataframes to quarterly frequency and to same length as all the other ones.'''
    # Setting the Date column
    df['Date'] = pd.to_datetime(df['observation_date'])
    
    # Selecting the needed data, dropping unneeded columns
    df = df[df["Date"].dt.month.isin(needed_months)].reset_index()
    df.drop(drop_list, axis=1, inplace=True)

    # Drop last observation optionally
    if drop_last_data == 'Y':
        df = df.iloc[:-1,:]
        
    return df

def visualize_shifts(df, date_column, series_1, series_2):
    # Extracting and preparing the DataFrames
    df1 = pd.DataFrame(df[[date_column, series_1]])
    df2 = pd.DataFrame(df[[date_column, series_2]])

    # Making sure dates are aligned
    df1[date_column] = pd.to_datetime(df1[date_column])
    df2[date_column] = pd.to_datetime(df2[date_column])

    # Merging data on dates
    merged_df = pd.merge(df1, df2, on = date_column)

    # Calculating correlation
    correlation = merged_df[series_1].corr(merged_df[series_2])
    print("Correlation Coefficient:", correlation)

    # Optional: Lag analysis
    lags = range(-10, 11)
    correlations = [merged_df[series_1].corr(merged_df[series_2].shift(lag)) for lag in lags]

    # Create a figure with two subplots
    fig, (ax1, ax2) = plt.subplots(2, 1, figsize=(12, 12))

    # Time Series Plot
    ax1.plot(merged_df[date_column], merged_df[series_1], label=f'Series 1 ({series_1})')
    ax1.plot(merged_df[date_column], merged_df[series_2], label=f'Series 2 ({series_2})')
    ax1.set_title('Time Series Plot')
    ax1.legend()
    ax1.set_xlabel('Date')
    ax1.set_ylabel('Normalized Values')

    # Correlation vs. Lag Plot
    ax2.plot(lags, correlations)
    ax2.set_title('Correlation vs. Lag')
    ax2.set_xlabel('Lag (Quarters)')
    ax2.set_ylabel('Correlation Coefficient')
    ax2.axvline(x=0, color='red', linestyle='--')  # Mark the zero lag line

    # Show the complete plot
    plt.tight_layout()
    plt.show()
    
def visualize_shifts_together4(df, date_column, series_1, series_2, series_3, series_4):
    # Extracting and preparing the DataFrames
    df1 = pd.DataFrame(df[[date_column, series_1]])
    df2 = pd.DataFrame(df[[date_column, series_2]])
    df3 = pd.DataFrame(df[[date_column, series_3]])
    df4 = pd.DataFrame(df[[date_column, series_4]])

    # Making sure dates are aligned
    df1[date_column] = pd.to_datetime(df1[date_column])
    df2[date_column] = pd.to_datetime(df2[date_column])
    df3[date_column] = pd.to_datetime(df3[date_column])
    df4[date_column] = pd.to_datetime(df4[date_column])

    # Merging data on dates
    merged_df = pd.merge(df1, df2, on = date_column, how = 'inner')
    merged_df = pd.merge(merged_df, df3, on = date_column, how = 'inner')
    merged_df = pd.merge(merged_df, df4, on = date_column, how = 'inner')

    # Calculating correlation
    correlation_12 = merged_df[series_1].corr(merged_df[series_2])
    correlation_13 = merged_df[series_1].corr(merged_df[series_3])
    correlation_42 = merged_df[series_4].corr(merged_df[series_2])
    correlation_43 = merged_df[series_4].corr(merged_df[series_3])
    print(f"Correlation Coefficient {series_1} vs {series_2}:", correlation_12)
    print(f"Correlation Coefficient {series_1} vs {series_3}:", correlation_13)
    print(f"Correlation Coefficient {series_4} vs {series_2}:", correlation_42)
    print(f"Correlation Coefficient {series_4} vs {series_3}:", correlation_43)
    
    # Optional: Lag analysis
    lags = range(-30, 31)
    correlations_12 = [merged_df[series_1].corr(merged_df[series_2].shift(lag)) for lag in lags]
    correlations_13 = [merged_df[series_1].corr(merged_df[series_3].shift(lag)) for lag in lags]
    correlations_42 = [merged_df[series_4].corr(merged_df[series_2].shift(lag)) for lag in lags]
    correlations_43 = [merged_df[series_4].corr(merged_df[series_3].shift(lag)) for lag in lags]

    # Create a figure with two subplots
    fig, (ax1, ax2) = plt.subplots(2, 1, figsize=(12, 14))

    # Time Series Plot
    ax1.plot(merged_df[date_column], merged_df[series_1], label=f'Series 1 ({series_1})')
    ax1.plot(merged_df[date_column], merged_df[series_2], label=f'Series 2 ({series_2})')
    ax1.plot(merged_df[date_column], merged_df[series_3], label=f'Series 3 ({series_3})')
    ax1.plot(merged_df[date_column], merged_df[series_4], label=f'Series 4 ({series_4})')
    ax1.set_title('Historical Yields for 1, 5, 10, and 30-Year Bonds (2013-2024)')
    ax1.legend()
    ax1.set_xlabel('Date')
    ax1.set_ylabel('Yields (%)')

    # Correlation vs. Lag Plot
    ax2.plot(lags, correlations_12, label=f'{series_1} vs {series_2}')
    ax2.plot(lags, correlations_13, label=f'{series_1} vs {series_3}')
    ax2.plot(lags, correlations_42, label=f'{series_4} vs {series_2}')
    ax2.plot(lags, correlations_43, label=f'{series_4} vs {series_3}')
    ax2.set_title('Correlation vs. Lag')
    ax2.legend()
    ax2.set_xlabel('Lag (Quarters)')
    ax2.set_ylabel('Correlation Coefficient')
    ax2.axvline(x=0, color='red', linestyle='--')  # Mark the zero lag line

    # Show the complete plot
    plt.subplots_adjust(hspace=0.3)
    plt.tight_layout()
    plt.show()

def visualize_shifts_together3(df, date_column, series_1, series_2, series_3):
    # Extracting and preparing the DataFrames
    df1 = pd.DataFrame(df[[date_column, series_1]])
    df2 = pd.DataFrame(df[[date_column, series_2]])
    df3 = pd.DataFrame(df[[date_column, series_3]])

    # Making sure dates are aligned
    df1[date_column] = pd.to_datetime(df1[date_column])
    df2[date_column] = pd.to_datetime(df2[date_column])
    df3[date_column] = pd.to_datetime(df3[date_column])

    # Merging data on dates
    merged_df = pd.merge(df1, df2, on = date_column, how = 'inner')
    merged_df = pd.merge(merged_df, df3, on = date_column, how = 'inner')

    # Calculating correlation
    correlation_12 = merged_df[series_2].corr(merged_df[series_3])
    correlation_13 = merged_df[series_1].corr(merged_df[series_3])
    print(f"Correlation Coefficient {series_2} vs {series_3}:", correlation_12)
    print(f"Correlation Coefficient {series_1} vs {series_3}:", correlation_13)
    
    # Optional: Lag analysis
    lags = range(-30, 31)
    correlations_12 = [merged_df[series_2].corr(merged_df[series_3].shift(lag)) for lag in lags]
    correlations_13 = [merged_df[series_1].corr(merged_df[series_3].shift(lag)) for lag in lags]

    # Create a figure with two subplots
    fig, (ax1, ax2) = plt.subplots(2, 1, figsize=(12, 14))

    # Time Series Plot
    ax1.plot(merged_df[date_column], merged_df[series_1], label=f'Series 1 ({series_1})')
    ax1.plot(merged_df[date_column], merged_df[series_2], label=f'Series 2 ({series_2})')
    ax1.plot(merged_df[date_column], merged_df[series_3], label=f'Series 3 ({series_3})')
    ax1.set_title('Normalized Trebds of 1-Year and 10-Year Rates with GDP Growth Rate 2013-2024')
    ax1.legend()
    ax1.set_xlabel('Date')
    ax1.set_ylabel('Normalized Values')

    # Correlation vs. Lag Plot
    ax2.plot(lags, correlations_12, label=f'{series_2} vs {series_3}')
    
    ax2.plot(lags, correlations_13, label=f'{series_1} vs {series_3}')
    ax2.set_title('Correlation vs. Lag')
    ax2.legend()
    ax2.set_xlabel('Lag (Quarters)')
    ax2.set_ylabel('Correlation Coefficient')
    ax2.axvline(x=0, color='red', linestyle='--')  # Mark the zero lag line

    # Show the complete plot
    plt.subplots_adjust(hspace=0.3)
    plt.tight_layout()
    plt.show()

# Define the Nelson-Siegel model function
def nelson_siegel_curve(t, beta0, beta1, beta2, tau):
    return beta0 + (beta1 + beta2) * ((1 - np.exp(-t / tau)) / (t / tau)) - beta2 * np.exp(-t / tau)

# Objective function to minimize (sum of squared errors)
def objective_function(params, t, rates):
    beta0, beta1, beta2, tau = params
    return np.sum((rates - nelson_siegel_curve(t, beta0, beta1, beta2, tau)) ** 2)

def visualize_shifts_together_double_y_axis(df, date_column, series_1, series_2, series_3, pvalue_series_1yr, pvalue_series_10yr):
    # Extracting and preparing the DataFrames
    df1 = pd.DataFrame(df[[date_column, series_1]])
    df2 = pd.DataFrame(df[[date_column, series_2]])
    df3 = pd.DataFrame(df[[date_column, series_3]])

    # Making sure dates are aligned
    df1[date_column] = pd.to_datetime(df1[date_column])
    df2[date_column] = pd.to_datetime(df2[date_column])
    df3[date_column] = pd.to_datetime(df3[date_column])

    # Merging data on dates
    merged_df = pd.merge(df1, df2, on=date_column, how='inner')
    merged_df = pd.merge(merged_df, df3, on=date_column, how='inner')

    # Calculating correlation
    correlation_12 = merged_df[series_2].corr(merged_df[series_3])
    correlation_13 = merged_df[series_1].corr(merged_df[series_3])
    print(f"Correlation Coefficient {series_2} vs {series_3}:", correlation_12)
    print(f"Correlation Coefficient {series_1} vs {series_3}:", correlation_13)
    
    # Optional: Lag analysis
    lags = range(-30, 31)
    correlations_12 = [merged_df[series_2].corr(merged_df[series_3].shift(lag)) for lag in lags]
    correlations_13 = [merged_df[series_1].corr(merged_df[series_3].shift(lag)) for lag in lags]

    # Create a figure with two subplots
    fig, (ax1, ax2) = plt.subplots(2, 1, figsize=(12, 14))

    # Time Series Plot
    ax1.plot(merged_df[date_column], merged_df[series_1], label=f'Series 1 ({series_1})')
    ax1.plot(merged_df[date_column], merged_df[series_2], label=f'Series 2 ({series_2})')
    ax1.plot(merged_df[date_column], merged_df[series_3], label=f'Series 3 ({series_3})')
    ax1.set_title('Normalized Trends of 1-Year and 10-Year Rates with S&P 500 Index (2013-2024)')
    ax1.legend()
    ax1.set_xlabel('Date')
    ax1.set_ylabel('Normalized Values')

    # Correlation vs. Lag Plot
    ax2.plot(lags, correlations_12, label=f'{series_2} vs {series_3}')
    ax2.plot(lags, correlations_13, label=f'{series_1} vs {series_3}')
    ax2.set_title('Correlation vs. Lag')
    ax2.legend(loc='upper left')
    ax2.set_xlabel('Lag (Quarters)')
    ax2.set_ylabel('Correlation Coefficient')
    ax2.axvline(x=0, color='red', linestyle='--')  # Mark the zero lag line

    # Create a second y-axis for the p-value series
    ax2_new = ax2.twinx()
    ax2_new.plot(lags, pvalue_series_1yr, label='1yr_pvalue', color='purple')
    ax2_new.plot(lags, pvalue_series_10yr, label='10yr_pvalue', color='orange')
    ax2_new.set_ylabel('P-values (p < 0.05 Significant)')
    ax2_new.legend(loc='upper right')

    # Show the complete plot
    plt.subplots_adjust(hspace=0.3)
    plt.tight_layout()
    plt.show()

def get_p_value_for_shifts(df, variable, tenors = [], min_lag = -30, max_lag = 31):

    p_values = [[] for i in range(len(tenors))]

    # Először a shiftelt oszlopok létrehozása, az összes shifttel
    for shift in list(range(min_lag, max_lag)):

        df[f"{variable}_shift_{shift}"] = df[variable].shift(shift)

        #Most modellezés minden shiftre
        X=df.dropna().copy()
        Y=X[tenors]
        X=X[f"{variable}_shift_{shift}"]

        # Transform the data with the 1/x connection we have noticed
        for column in Y.columns:
            Y[f"{column}_1perX"] = 1/(1 + Y[column])

        #We need to add a constant value   
        X = sm.add_constant(X)

        #Let us finally run the OLS
        results = {}
        for i in range(Y.shape[1]):  # Iterate over each column in Y.
            model = sm.OLS(Y.iloc[:, i], X.astype(float)).fit()  # Fit model for Y's ith column.
            results[Y.columns[i]] = model  # Store the summary with the column name as key.

        for i in range(len(tenors)):
            all_variable_p_values = results[f'{tenors[i]}_1perX'].pvalues
            p_values[i].append(all_variable_p_values[f"{variable}_shift_{shift}"])

    return p_values

def objective_function(params, t, rates):
    beta0, beta1, beta2, lambd = params
    return np.sum((rates - nelson_siegel_curve(t, beta0, beta1, beta2, lambd)) ** 2)

def nelson_siegel_curve(t, beta0, beta1, beta2, lambd):
    return beta0 + (beta1 + beta2) * ((1 - np.exp(-t / lambd)) / (t / lambd)) - beta2 * np.exp(-t / lambd)


