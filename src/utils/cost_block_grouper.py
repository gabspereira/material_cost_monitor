import pandas as pd



def Grouper(df):
    # call df dataframe
    df = df
    print('Call df dataframe.')

    # Select only useful columns
    column_ids = [0,2,3,4,5]
    df = df.iloc[:,column_ids]
    print('Select only useful columns.')

    # Remove unnecessary points into level column
    df.iloc[:,0] = df.iloc[:,0].str.replace('.', '')
    print('Remove unnecessary points into level column.')

    # Group Cost Blocks
    df['cost_block'] = 0

    for i in range(len(df)):
        if df['Component number'].iloc[i].find('A7ET') == 0 and df['Component number'].iloc[i].find('A7ET0') != 0:
            memory = df['Component number'].iloc[i]
            df['cost_block'].iloc[i] = memory
        else:
            df['cost_block'].iloc[i] = memory

    print('Group Cost Blocks')

    return df