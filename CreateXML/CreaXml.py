import xml.etree.ElementTree as gfg
from importaLotti import importaLottiPannelli
import pymssql
conn=pymssql.connect(server='192.168.100.***',user='***',password='****',database='******')
cursor=conn.cursor()
importaLottiPannelli(lottoIn,LottoFin)
cursor.execute("select CodiceArticolo from TE_LOTTI_SEZIONALI where Lotto='"+lottoIn+"'")
ArticoliInLotto=cursor.fetchall()
def GenerateXML(fileName):
    root=gfg.Element("Lotto")
    root.text=lottoIn
    
    for rows in range(len(ArticoliInLotto)):
        
        cursor.execute("SELECT lot.*,rad.CodiceCaratteristica,rad.Valore,rad.fStringa,rad.valorestringa FROM TE_LOTTI_SEZIONALI lot JOIN R_ARTICOLO_DNA rad ON rad.CodiceArticolo=lot.CodiceArticolo WHERE lot.Lotto='"+lottoIn+"' and lot.CodiceArticolo="+ArticoliInLotto[rows][0]+"")
        for riga in cursor:
            m1=gfg.Element("OrdProd")
            m1.text=riga[3]
            root.append(m1)

            b1=gfg.SubElement(m1,"OrdCli")
            b1.text=riga[2]+"-"+riga[4]


            b11=gfg.SubElement(b1,"Pannello")
            b11.text="TOP"
            b11=gfg.SubElement(b1,"L")
            cursor.execute("select ValoreCar from CAR_ACQ_CONF_COMM where CodiceCaratteristicaAcq='SZ026' AND Codice='"+riga[0]+"'")
            L=cursor.fetchone()
            b11.text=str(int(L[0]))
            b11=gfg.SubElement(b1,"H")
            cursor.execute("select ValoreCar from CAR_ACQ_CONF_COMM where CodiceCaratteristicaAcq='SZ028' AND Codice='"+riga[0]+"'")
            H=cursor.fetchone()
            b11.text=str(int(H[0]))

            cursor.execute("select ValoreCar from CAR_ACQ_CONF_COMM where CodiceCaratteristicaAcq='SZ027' AND Codice='"+riga[0]+"'")
            NPINT=cursor.fetchone()
            nPan=int(NPINT[0])
            for i in range(nPan):
                b11=gfg.SubElement(b1,"Pannello")
                b11.text="PANNELLO"+str(nPan-i)
                b11=gfg.SubElement(b1,"L")
                b11.text=str(int(L[0]))
                b11=gfg.SubElement(b1,"H")
                b11.text="500"
                #estrazione dati oblo
                b11=gfg.SubElement(b1,"Feature")
                cursor.execute("SELECT ValoreCar from CAR_PROPRIE_CONF_COMM cpcc where Codice ='"+riga[0]+"' and CodCarPropria='SZ25"+str(nPan-i-1)+"' and ValoreCar >0")
                oblo=cursor.fetchone()
                if oblo is not None:
                    b11=gfg.SubElement(b1,"Oblo")
                    b11.text=str(int(oblo[0]))

                    #estrazione dimensioni
                    cursor.execute("SELECT CodCarPropria ,ValoreCar from CAR_PROPRIE_CONF_COMM cpcc where Codice ='"+riga[0]+"' and CodCarPropria in('SZ450','SZ460','SZ470') and ValoreCar >0")
                    dimOblo=cursor.fetchall()
                    for dim in range(len(dimOblo)):
                        if dimOblo[dim][0]=='SZ450':
                            b11=gfg.SubElement(b1,"L_Oblo")
                            b11.text=str(int(dimOblo[dim][1]))
                        elif dimOblo[dim][0]=='SZ460':
                            b11=gfg.SubElement(b1,"H_Oblo")
                            b11.text=str(int(dimOblo[dim][1]))
                        elif dimOblo[dim][0]=='SZ470':
                            b11=gfg.SubElement(b1,"D_Oblo")
                            b11.text=str(int(dimOblo[dim][1]))



                
    tree=gfg.ElementTree(root)

    with open(fileName,'wb') as files:
        tree.write(files)

#Driver code
if __name__=='__main__':
    GenerateXML("C://Users//fabrizio_beggiato//PycharmProjects//PannelliXML//Catalog.xml")
