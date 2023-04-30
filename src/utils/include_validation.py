"""
The functions below intends to classify if the material
should be purchase or not, based on how the materials had been created.

Paramenters
-----------
type = ['F', 'E', 'X']
bulk = [nan, 'X']
include = [1, 0]

"""
import pandas as pd


""" 
In order to define include or not inclue values
a join between TABLE MARC and BOM is required
"""





def Purchase(df):
    # call bill_of_material dataframe
    df = df
    
    # call table_marc dataframe
    table_marc = pd.read_excel('C:/Temp//data/SAP/TABLE_MARC.xlsx')
    print('Table Marc loaded.')
    
    # join bill of material and marc table
    df = df.merge(table_marc[['Material Number', 'Material Description', 'Procurement type', 'Bulk material']], \
        left_on='Component number', \
        right_on='Material Number', \
        how='left'
        )
    print('Procurement type, Bulk material include into the table')

    # organize resulted table
    df = df[['Explosion level', \
            'cost_block', \
            'Component number', \
            'Material Description', \
            'Comp. Qty (CUn)', \
            'Component unit', \
            'Procurement type', \
            'Bulk material'
        ]]
    print('Columns reduced.')

    # rename resulted table
    df.rename(columns={'Explosion level':'level', \
                        'Component number':'material', \
                        'Procurement type':'type', \
                        'Bulk material':'bulk', \
                        'Comp. Qty (CUn)':'qty', \
                        'Component unit':'unit', \
                        'Material Description':'description'}, \
                        inplace=True)
    print('Columns renamed.')

    # fill NAN values
    df['bulk'].fillna('', inplace=True)
    print('NaN values filled.')


    # Convert level into numeric
    df['level'] = df['level'].astype(int)
    print('Numeric values converted.')


    df['include']=''
    lvl_list = len(df['level'].unique())
    i=1

    # Limitar o agrupamento do bloco de custo (ex. do level 1 até 9)
    while i <= lvl_list:
        try:

            # inicio da iteração  
            for j in range(len(df)):

                # Se ex. level=1, type=F, bulk='', include=''; então definir include=1
                if df['level'].iloc[j] == i and df['type'].iloc[j] == 'F' and df['bulk'].iloc[j] == '' and df['include'].iloc[j] == '':
                    df.at[j,'include'] = 1
                    
                    # Zerar todos os itens filhos do mesmo bloco, uma vez que o item pai, include=1
                    k = j + 1
                    try:
                        while df['level'].iloc[k] != i and df['include'].iloc[k] == '':
                            df.at[k,'include'] = 0
                            k=k+1
                    except: pass
            i=i+1
        except: pass

    # Caso sobre algum item com include='', então include=0
    for z in range(len(df)):
        if df['include'].iloc[z] == '':
            df.at[z, 'include'] = 0
    
    print('Include_validation function finalized.')
    return df