def importaLottiPannelli(lotin,lotfin):
    
    #connessione galileo
    import pyodbc
    conn=pyodbc.connect(
        driver='{iSeries Access ODBC Driver}',
        system='192.168.100.4',
        uid='GALILEO80',
        pwd='GALILEO80')
    cursor=conn.cursor()

    #connessione tce
    import pymssql
    conn1=pymssql.connect(server='192.168.100.204',user='sa',password='trinity',database='TCEBALLAN')
    cursor1=conn1.cursor()

    #lotin='2S176'
    #lotfin='2S176'
    cursor.execute("SELECT CDARPO,QORDPO,RIFEPO,ORPRPO,CLIEPO,LOTTPO from BAL80DAT.PMORD00f where LOTTPO>='"+lotin+"' and LOTTPO<='"+lotfin+"' and cdarpo like '8%'")
    for rows in cursor:
        codice=rows[0][0:13]
        riferimento=rows[2][3:10]
        quantita=int(rows[1])
        ord_prod=rows[3]
        cliente=rows[4]
        n_lotto=rows[5]
        #scrivo in tce
        cursor1.execute("INSERT INTO TE_LOTTI_SEZIONALI (CodiceArticolo,Quantita,Riferimento,OrdProd,Cliente,Lotto) VALUES ('"+codice+"','"+str(int(quantita))+"','"+riferimento+"','"+ord_prod+"','"+cliente+"','"+n_lotto+"')")
        conn1.commit()

    conn.close()
    return (codice,quantita,riferimento,ord_prod,cliente,n_lotto)

