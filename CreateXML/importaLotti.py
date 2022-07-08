def importaLottiPannelli(lotin,lotfin):
    
    #connessione Galileo
    import pyodbc
    conn=pyodbc.connect(
        driver='{iSeries Access ODBC Driver}',
        system='192.168.100.***',
        uid='****',
        pwd='****')
    cursor=conn.cursor()

    #connessione MS-SQL
    import pymssql
    conn1=pymssql.connect(
        server='192.168.100.***',
        user='***',
        password='*****',
        database='****')
    cursor1=conn1.cursor()

    cursor.execute("SELECT CDARPO,QORDPO,RIFEPO,ORPRPO,CLIEPO,LOTTPO from BAL80DAT.PMORD00f where LOTTPO>='"+lotin+"' and LOTTPO<='"+lotfin+"' and cdarpo like '8%'")
    for rows in cursor:
        codice=rows[0][0:13]
        riferimento=rows[2][3:10]
        quantita=int(rows[1])
        ord_prod=rows[3]
        cliente=rows[4]
        n_lotto=rows[5]
        #scrivo in MS-SQL
        cursor1.execute("INSERT INTO TE_LOTTI_SEZIONALI (CodiceArticolo,Quantita,Riferimento,OrdProd,Cliente,Lotto) VALUES ('"+codice+"','"+str(int(quantita))+"','"+riferimento+"','"+ord_prod+"','"+cliente+"','"+n_lotto+"')")
        conn1.commit()

    conn.close()
    return (codice,quantita,riferimento,ord_prod,cliente,n_lotto)

