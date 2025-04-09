import pandas as pd

def run_conciliacion(df_RB, df_LB, Debe_RB, Haber_RB, Debe_LB, Haber_LB):
    # Evitamos modificar los originales
    df1 = df_RB.copy()
    df2 = df_LB.copy()
    df2 = df2.drop(index=0)  # Si hay encabezado duplicado, se elimina

    # ||||| PASO 1 |||||
    freq_df2 = df2[Haber_LB].value_counts().to_dict()
    en_df2 = []

    for valor in df1[Debe_RB]:
        if freq_df2.get(valor, 0) > 0:
            en_df2.append(True)
            freq_df2[valor] -= 1
        else:
            en_df2.append(False)

    df1['En_df2'] = en_df2
    #df1p1_filtrado = df1[df1['En_df2'] == False]
    #df1p1final = df1p1_filtrado[[col for col in df1.columns if 'clasificacion' in col.lower() or 'debit' in col.lower()]]
    dfpaso1 = df1
    
    # ||||| PASO 2 |||||
    freq_df1 = df1[Debe_RB].value_counts().to_dict()
    en_df1 = []

    for valor in df2[Haber_LB]:
        if freq_df1.get(valor, 0) > 0:
            en_df1.append(True)
            freq_df1[valor] -= 1
        else:
            en_df1.append(False)

    df2['En_df1'] = en_df1
    #df2p2_filtrado = df2[df2['En_df1'] == False]
    #df2p2final = df2p2_filtrado[[col for col in df2.columns if 'clasificacion' in col.lower() or 'haber' in col.lower()]]
    dfpaso2 = df2

    # ||||| PASO 3 |||||
    freq_df2 = df2[Debe_LB].value_counts().to_dict()
    en_df2 = []

    for valor in df1[Haber_RB]:
        if freq_df2.get(valor, 0) > 0:
            en_df2.append(True)
            freq_df2[valor] -= 1
        else:
            en_df2.append(False)

    df1['En_df2_p3'] = en_df2
    #df1p3_filtrado = df1[df1['En_df2_p3'] == False]
    #df1p3final = df1p3_filtrado[[col for col in df1.columns if 'clasificacion' in col.lower() or 'credito' in col.lower()]]
    dfpaso3 = df1

    # ||||| PASO 4 |||||
    freq_df1 = df1[Haber_RB].value_counts().to_dict()
    en_df1 = []

    for valor in df2[Debe_LB]:
        if freq_df1.get(valor, 0) > 0:
            en_df1.append(True)
            freq_df1[valor] -= 1
        else:
            en_df1.append(False)

    df2['En_df1_p4'] = en_df1
    #df2p4_filtrado = df2[df2['En_df1_p4'] == False]
    #df2p4final = df2p4_filtrado[[col for col in df2.columns if 'clasificacion' in col.lower() or 'debe' in col.lower()]]
    dfpaso4 = df2

    # ||||| RESUMEN |||||
    # Columnas para sumar
    col_debito = next((c for c in df1.columns if 'debit' in c.lower()), None)
    col_credito = next((c for c in df1.columns if 'credit' in c.lower()), None)
    col_haber = next((c for c in df2.columns if 'haber' in c.lower()), None)
    col_debe = next((c for c in df2.columns if 'debe' in c.lower()), None)
    col_saldo_RB = next((c for c in df1.columns if 'saldo' in c.lower()), None)
    col_saldo_LB = next((c for c in df2.columns if 'saldo' in c.lower()), None)

    saldofinalRB = float(df1[col_saldo_RB].iloc[-1])
    saldofinalLB = float(df2[col_saldo_LB].iloc[-1])

    SumPaso1 = dfpaso1[col_debito].sum() if col_debito else 0
    SumPaso2 = dfpaso2[col_haber].sum() if col_haber else 0
    SumPaso3 = dfpaso3[col_credito].sum() if col_credito else 0
    SumPaso4 = dfpaso4[col_debe].sum() if col_debe else 0

    SaldoLB = saldofinalRB + SumPaso1 - SumPaso2 - SumPaso3 + SumPaso4
    difsaldos = SaldoLB - saldofinalLB

    resumen_data = {
        'Pasos': ['Suma Paso 1', 'Suma Paso 2', 'Suma Paso 3','Suma Paso 4','Saldo Final RB','Saldo Final LB','Saldo pasos(+y-)','Diferencia Saldos'],
        'Resultado': [SumPaso1, SumPaso2, SumPaso3, SumPaso4, saldofinalRB, saldofinalLB, SaldoLB, difsaldos]
    }

    df3 = pd.DataFrame(resumen_data)

    return dfpaso1, dfpaso2, dfpaso3, dfpaso4, df3
