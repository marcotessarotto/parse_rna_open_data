import xml.sax

search_results = {}

DEBUG_MSG = False

results_counter = 0


class RNAHandler(xml.sax.ContentHandler):
    def __init__(self):
        self.CurrentData = ""

        self.CAR = ""
        self.SOGGETTO_CONCEDENTE = ""
        self.DENOMINAZIONE_UFF_GESTORE = ""
        self.COD_UFF_GESTORE = ""
        self.BASE_GIURIDICA_NAZIONALE = ""
        self.TITOLO_PROGETTO = ""
        self.COR = ""
        self.REGIONE_BENEFICIARIO = ""
        self.CUP = ""

    # Call when an element starts
    def startElement(self, tag, attributes):
        self.CurrentData = tag

    # Call when an elements ends
    def endElement(self, tag):

        if tag == "AIUTO":

            ok = True if "REGIONE AUTONOMA FRIULI VENEZIA GIULIA - DIREZIONE CENTRALE LAVORO, FORMAZIONE, ISTRUZIONE E FAMIGLIA" in self.SOGGETTO_CONCEDENTE else False

            # ok = True if "REGIONE AUTONOMA FRIULI VENEZIA GIULIA" in self.SOGGETTO_CONCEDENTE else False

            if ok:
                if DEBUG_MSG:
                    print(f"\nCAR: {self.CAR}")
                    print(f"DENOMINAZIONE_UFF_GESTORE: {self.DENOMINAZIONE_UFF_GESTORE}")
                    print(f"COD_UFF_GESTORE: {self.COD_UFF_GESTORE}")
                    print(f"SOGGETTO_CONCEDENTE: {self.SOGGETTO_CONCEDENTE}")
                    print(f"BASE_GIURIDICA_NAZIONALE: {self.BASE_GIURIDICA_NAZIONALE}")
                    print(f"TITOLO_PROGETTO: {self.TITOLO_PROGETTO}")
                    print(f"COR: {self.COR}")
                    print(f"REGIONE_BENEFICIARIO: {self.REGIONE_BENEFICIARIO}")
                    print(f"CUP: {self.CUP}")

                d = {"CAR": self.CAR,
                     "DENOMINAZIONE_UFF_GESTORE": self.DENOMINAZIONE_UFF_GESTORE,
                     "COD_UFF_GESTORE": self.COD_UFF_GESTORE,
                     "SOGGETTO_CONCEDENTE": self.SOGGETTO_CONCEDENTE,
                     "BASE_GIURIDICA_NAZIONALE": self.BASE_GIURIDICA_NAZIONALE,
                     "TITOLO_PROGETTO": self.TITOLO_PROGETTO,
                     "COR": self.COR,
                     "REGIONE_BENEFICIARIO": self.REGIONE_BENEFICIARIO,
                     "CUP": self.CUP
                     }

                global results_counter

                search_results[results_counter] = d

                results_counter = results_counter + 1

            # print("END***")
            self.CAR = ""
            self.DENOMINAZIONE_UFF_GESTORE = ""
            self.COD_UFF_GESTORE = ""
            self.SOGGETTO_CONCEDENTE = ""
            self.BASE_GIURIDICA_NAZIONALE = ""
            self.TITOLO_PROGETTO = ""
            self.COR = ""
            self.REGIONE_BENEFICIARIO = ""
            self.CUP = ""

        self.CurrentData = ""

    # Call when a character is read
    def characters(self, content):
        # print(self.CurrentData)

        if self.CurrentData == "CAR":
            self.CAR = content
        elif self.CurrentData == "SOGGETTO_CONCEDENTE":
            self.SOGGETTO_CONCEDENTE = content
        elif self.CurrentData == "DENOMINAZIONE_UFF_GESTORE":
            self.DENOMINAZIONE_UFF_GESTORE = content
        elif self.CurrentData == "COD_UFF_GESTORE":
            self.COD_UFF_GESTORE = content
        elif self.CurrentData == "BASE_GIURIDICA_NAZIONALE":
            self.BASE_GIURIDICA_NAZIONALE = content
        elif self.CurrentData == "TITOLO_PROGETTO":
            self.TITOLO_PROGETTO = content
        elif self.CurrentData == "COR":
            self.COR = content
        elif self.CurrentData == "REGIONE_BENEFICIARIO":
            self.REGIONE_BENEFICIARIO = content
        elif self.CurrentData == "CUP":
            self.CUP = content


if __name__ == "__main__":
    # create an XMLReader
    parser = xml.sax.make_parser()
    # turn off namepsaces
    parser.setFeature(xml.sax.handler.feature_namespaces, 0)

    # override the default ContextHandler
    Handler = RNAHandler()
    parser.setContentHandler(Handler)

    files = (
        "C:\\Users\\140536\\Downloads\\OpenData_Aiuti_2020_01.xml",
        "C:\\Users\\140536\\Downloads\\OpenData_Aiuti_2020_02.xml",
        "C:\\Users\\140536\\Downloads\\OpenData_Aiuti_2020_03.xml",
        "C:\\Users\\140536\\Downloads\\OpenData_Aiuti_2020_04.xml",
        "C:\\Users\\140536\\Downloads\\OpenData_Aiuti_2020_05.xml",
        "C:\\Users\\140536\\Downloads\\OpenData_Aiuti_2020_06.xml",
        "C:\\Users\\140536\\Downloads\\OpenData_Aiuti_2020_07.xml",
        "C:\\Users\\140536\\Downloads\\OpenData_Aiuti_2020_08.xml",
        "C:\\Users\\140536\\Downloads\\OpenData_Aiuti_2020_09.xml",
        "C:\\Users\\140536\\Downloads\\OpenData_Aiuti_2020_10.xml",
        "C:\\Users\\140536\\Downloads\\OpenData_Aiuti_2020_11.xml",
        "C:\\Users\\140536\\Downloads\\OpenData_Aiuti_2020_12.xml",
        "C:\\Users\\140536\\Downloads\\OpenData_Aiuti_2020_12_002.xml",
    )

    for f in files:
        print(f"processing {f}...")
        parser.parse(f)
        print()
        print(len(search_results))
        print()


    print()
    print(len(search_results))
    print()

    import xlsxwriter

    # Create a workbook and add a worksheet.
    workbook = xlsxwriter.Workbook('results.xlsx')
    worksheet = workbook.add_worksheet()

    # Start from the first cell. Rows and columns are zero indexed.
    row = 0
    col = 0

    worksheet.write(row, col + 0, "CAR")
    worksheet.write(row, col + 1, "DENOMINAZIONE_UFF_GESTORE")
    worksheet.write(row, col + 2, "COD_UFF_GESTORE")
    worksheet.write(row, col + 3, "SOGGETTO_CONCEDENTE")
    worksheet.write(row, col + 4, "BASE_GIURIDICA_NAZIONALE")
    worksheet.write(row, col + 5, "TITOLO_PROGETTO")
    worksheet.write(row, col + 6, "COR")
    worksheet.write(row, col + 7, "REGIONE_BENEFICIARIO")
    worksheet.write(row, col + 8, "CUP")

    row += 1

    # Iterate over the data and write it out row by row.
    for cup, d in search_results.items():
        worksheet.write(row, col + 0, d["CAR"])
        worksheet.write(row, col + 1, d["DENOMINAZIONE_UFF_GESTORE"])
        worksheet.write(row, col + 2, d["COD_UFF_GESTORE"])
        worksheet.write(row, col + 3, d["SOGGETTO_CONCEDENTE"])
        worksheet.write(row, col + 4, d["BASE_GIURIDICA_NAZIONALE"])
        worksheet.write(row, col + 5, d["TITOLO_PROGETTO"])
        worksheet.write(row, col + 6, d["COR"])
        worksheet.write(row, col + 7, d["REGIONE_BENEFICIARIO"])
        worksheet.write(row, col + 8, d["CUP"])
        row += 1

    # Write a total using a formula.
    # worksheet.write(row, 0, 'Total')
    # worksheet.write(row, 1, '=SUM(B1:B4)')

    workbook.close()

"""
<?xml version="1.0" encoding="UTF-8"?>
<AIUTO xmlns="http://www.rna.it/RNA_aiuto/schema">
   <CAR>802</CAR>
   <TITOLO_MISURA>FVG-AIUTI ALL'OCCUPAZIONE DI SOGGETTI DISABILI</TITOLO_MISURA>
   <DES_TIPO_MISURA>Regime di aiuti</DES_TIPO_MISURA>
   <COD_CE_MISURA>SA.46707 (2016/X)</COD_CE_MISURA>
   <BASE_GIURIDICA_NAZIONALE>REGIONE FVG - Legge regionale 9 agosto 2005 n. 18 Norme regionali per l'occupazione, la tutela e la qualità del lavoro</BASE_GIURIDICA_NAZIONALE>
   <LINK_TESTO_INTEGRALE_MISURA>http://www.regione.fvg.it/rafvg/cms/RAFVG/formazione-lavoro/lavoro/FOGLIA117/</LINK_TESTO_INTEGRALE_MISURA>
   <COD_UFF_GESTORE>LAV_MIRATO</COD_UFF_GESTORE>
   <DENOMINAZIONE_UFF_GESTORE>PO COLLOCAMENTO MIRATO</DENOMINAZIONE_UFF_GESTORE>
   <SOGGETTO_CONCEDENTE>REGIONE AUTONOMA FRIULI VENEZIA GIULIA - DIREZIONE CENTRALE LAVORO, FORMAZIONE, ISTRUZIONE E FAMIGLIA</SOGGETTO_CONCEDENTE>
   <COR>1623443</COR>
   <TITOLO_PROGETTO>FONDO REGIPONALE DISABILI</TITOLO_PROGETTO>
   <DESCRIZIONE_PROGETTO>ASSUZNIONE TI DI UN LAVORATORE CON DISABILITA'</DESCRIZIONE_PROGETTO>
   <DATA_CONCESSIONE>2020-02-03+01:00</DATA_CONCESSIONE>
   <CUP>D48E20000020002</CUP>
   <ATTO_CONCESSIONE>861</ATTO_CONCESSIONE>
   <DENOMINAZIONE_BENEFICIARIO>STREAM YACHTS S.R.L.</DENOMINAZIONE_BENEFICIARIO>
   <CODICE_FISCALE_BENEFICIARIO>04145400265</CODICE_FISCALE_BENEFICIARIO>
   <DES_TIPO_BENEFICIARIO>PMI</DES_TIPO_BENEFICIARIO>
   <REGIONE_BENEFICIARIO>Friuli-Venezia Giulia</REGIONE_BENEFICIARIO>
   <COMPONENTI_AIUTO>
      <COMPONENTE_AIUTO>
         <ID_COMPONENTE_AIUTO>1879214</ID_COMPONENTE_AIUTO>
         <COD_PROCEDIMENTO>3</COD_PROCEDIMENTO>
         <DES_PROCEDIMENTO>Esenzione</DES_PROCEDIMENTO>
         <COD_REGOLAMENTO>CE651/2014</COD_REGOLAMENTO>
         <DES_REGOLAMENTO>Reg. CE 651/2014 esenzione generale per categoria (GBER)</DES_REGOLAMENTO>
         <COD_OBIETTIVO>502004</COD_OBIETTIVO>
         <DES_OBIETTIVO>Aiuti a favore di lavoratori svantaggiati e di lavoratori con disabilità | Aiuti all'occupazione di lavoratori con disabilita sotto forma di integrazioni salariali (art. 33)</DES_OBIETTIVO>
         <SETTORE_ATTIVITA>C.30.1</SETTORE_ATTIVITA>
         <STRUMENTI_AIUTO>
            <STRUMENTO_AIUTO>
               <COD_STRUMENTO>37</COD_STRUMENTO>
               <DES_STRUMENTO>Sovvenzione/Contributo in conto interessi</DES_STRUMENTO>
               <ELEMENTO_DI_AIUTO>8498.44</ELEMENTO_DI_AIUTO>
               <IMPORTO_NOMINALE>8498.44</IMPORTO_NOMINALE>
            </STRUMENTO_AIUTO>
         </STRUMENTI_AIUTO>
      </COMPONENTE_AIUTO>
   </COMPONENTI_AIUTO>
</AIUTO>
"""
