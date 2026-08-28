from pathlib import Path
from bs4 import BeautifulSoup
import pandas as pd
import re
import unicodedata
pd.set_option("display.max_colwidth", None)
pd.set_option("display.max_rows", 200)

# HTML-Datei auswählen
def search_decisions(soup):

    entscheid = "andere"
    franc=0
    deutsch=0
    ital=0
    unclear=0

    def entscheid_kategorisieren(schluss_text):

        text = schluss_text.lower()

        for kategorie, muster_liste in entscheidungsmuster.items():
            for muster in muster_liste:
                if re.search(muster, text):
                    return kategorie

        return "andere"


    french = [
        "par ces motifs",
        "le tribunal fédéral prononce",
         "p ar ces motifs, le juge instructeur ordonne :",
        "pour ces motifs, le tribunal fédéral suisse constate:",
        "p ar ces motifs, le juge présidant prononce :",
        "p ar ces motifs, la présidente prononce :",
        "p ar ces motifs, le président prononce :",
    ]
    
    german = [
        "demnach erkennt",
        "erkennt die präsidentin:",
        "erkennt der präsident:",
        "demnach verfügt der einzelrichter",
        "demnach verfügt der präsident",
        "erkennt der einzelrichter",
        "demnach verfügt der instruktionsrichter",
        "demnach verfügt die instruktionsrichterin",
        "erkennt die einzelrichterin",
        "verfügt die präsidentin",      
        "verfügt der präsident",
        "verfügt die einzelrichterin",
        "demnach stellt das schweizerische bundesgericht fest",
        "das bundesgericht erkennt",
        "demnach verfügt das präsidierende mitglied",
    
         "erkennt das präsidierende mitglied :",
        "verfügt der instruktionsrichter:",
        "verfügt das präsidierende mitglied :",
        "demnach erkenn t das bundesgericht:",
        "demnach wird verfügt und beschlossen:",
        "demnach stellt das schweizerischen bundesgericht fest:",
        "demnach verfügt und erkennt das präsidierende mitglied :",
        "die präsidentin verfügt:",
        "demnach kennt das bundesgericht:",
    
        "demnach verfügt das bundesgericht:",
        "verfügt das präsidierende mitglied:",
        "demnach entscheidet das bundesgericht:",
        "erkennt das bundesgericht:",
        "erkennt das präsidierende mitglied:",
        "erkennt das präsidierende mitglied,",
        "demnach erkennt das bundesgericht:",
        "verfügt der einzelrichter:",
        "demnach verfügt das bundesgericht",
        "verfügt das bundesgericht",
        "demnach verfügt das präsidierende Mitglied",
         "erkennt das präsidierende mitglied :",
        "beschliesst das bundesgericht:",
       
        
    ]
    
    italy = [
        "per questi motivi",
         "il tribunale federale pronuncia:",
    ]


    muster = "|".join(
        re.escape(formel)
        for formel in french + german + italy
    )

    muster = muster + r"|^\s*ordonne\s*:\s*$"



    entscheidungsmuster = {
    
        
        "teilweise gutgeheissen": [
            r"wird teilweise gutgeheissen",
            r"werden teilweise gutgeheissen",
            r"le recours est partiellement admis",
            r"la recours est partiellement admis",
            r"il ricorso è parzialmente accolto",
            r"die beschwerde wird, soweit .*? gutgeheissen",
            r"il ricorso è parzialmente ammesso",
            r"die beschwerde wird, soweit darauf eingetreten wird, teilweise gutgeheissen",
            r"wird.*?teilweise gutgeheissen",
            r"le recours.*?est partiellement admis",
             r"les recours.*?sont partiellement admis",
            r"i ricorsi.*?sono parzialmente accolti",
            r"il ricorso.*?è parzialmente accolto",
            r"in teilweiser gutheissung der beschwerde",
             r"die beschwerde wird teilweise.*?gutgeheissen",
        r"le recours.*?et le recours.*?sont partiellement admis",
        ],
    
        "gutgeheissen": [
            r"la requête.*?doit être admise",
            r"wird gutgeheissen",
            r"werden gutgeheissen",
            r"le recours est admis",
            r"la demande.*?est admise",
            r"les .*?demandes.*?sont admises",
            r"in gutheissung der beschwerde",
            r"la requête.*?est admise",
            r"les demandes.*?sont admises",
            r"la recours est admis",
            r"il ricorso è accolto",
            r"il ricorso è ammesso",
            r"le recours.*?est admis",
             r"les recours.*?sont admis",
            r"i ricorsi.*?sono accolti",
            r"il ricorso.*?è accolto",
            r"die beschwerden.*?werden gutgeheissen",
            r"die beschwerde wird(?:\s+\S+){0,6}\s+gutgeheissen",
             r"wird die beschwerde gutgeheissen",
            r"in dahingehender gutheissung der beschwerde",
        ],
    
        "abgeschrieben": [
            r"wird .*? abgeschrieben",
            r"werden .*? abgeschrieben",
            r"wird als erledigt abgeschrieben",
            r"wird infolge rückzugs .*? abgeschrieben",
            r"est rayée du rôle",
            r"est rayé du rôle",
            r"sont rayées du rôle",
            r"il est pris acte du retrait",
            r"divenuta priva d'oggetto",
            r"è stralciata dai ruoli",
             r"la causa.*?è stralciata dal ruolo",
            r"die verfahren.*?werden abgeschrieben",
            r"das verfahren wird abgeschrieben",
            r"die verfahren werden abgeschrieben",
            r"das verfahren.*?wird abgeschrieben",
            r"est radiée du rôle",
        ],
    
        "nicht eingetreten": [
            r"wird nicht eingetreten",
            r"l'opposition.*?est irrecevable",
            r"wird darauf nicht eingetreten",
            r"auf die beschwerde wird nicht eingetreten",
            r"le recours est(?: manifestement)? irrecevable",
            r"le recours est irrecevable",
            r"la domanda.*?è inammissibile",
            r"le recours est(?: déclaré| manifestement)? irrecevable",
            r"il ricorso.*?è inammissibile",
            r"le recours.*?et le recours.*?sont irrecevables",
            r"la recours est irrecevable",
            r"il ricorso è inammissibile",
            r"il ricorso è irricevibile",
            r"la domanda di revisione.*?è inammissibile",
            r"le recours.*?est irrecevable",
            r"la demande.*?et la requête.*?sont irrecevables",
            r"il n'est pas entré en matière",
            r"les recours.*?sont irrecevables",
            r"la demande.*?est irrecevable",
            r"il ricorso.*?è inammissibile",
            r"la requête.*?est irrecevable",
            r"le mémoire.*?est irrecevable",
            r"l'écriture.*?est irrecevable",
            r"i ricorsi.*?sono inammissibili",
            r"il ricorso.*?è inammissibile",
            r"auf die beschwerden.*?wird nicht eingetreten",
            r"la domanda.*?è inammissibile",
            r"le courrier.*?est irrecevable.*?constitue un recours",
            r"soweit.*?wird auf die beschwerde nicht eingetreten",
             r"les requêtes.*?sont irrecevables",
            r"la dénonciation, traitée comme recours.*?est irrecevable",
           
        ],
    
        "abgewiesen": [
            r"wird abgewiesen",
            r"werden abgewiesen",
            r"le recours est rejeté",
            r"la recours est rejetée",
            r"il ricorso è respinto",
            r"il ricorso è rigettato",
            r"le recours.*?doit être rejeté",
            r"die beschwerde ist abzuweisen, soweit auf sie einzutreten ist",
            r"die beschwerden sind abzuweisen, soweit auf sie einzutreten ist",
            r"la demande de révision est rejetée",
            r"le recours.*?est rejeté",
            r"la demande de restitution du délai.*?est rejetée",
            r"les recours.*?sont rejetés",
            r"la demande.*?est rejetée",
            r"la requête.*?est rejetée",
            r"i ricorsi.*?sono respinti",
            r"la domanda di revisione.*?è respinta",
            r"il ricorso.*?è respinto",
            r"die beschwerden.*?werden abgewiesen",
            r"la domanda di revisione.*?è respinta",
            r"die beschwerde.*?wird.*?abgewiesen.*?soweit darauf eingetreten wird",
            r"die beschwerden.*?werden.*?abgewiesen.*?soweit darauf eingetreten wird",
            r"les .*?recours.*?sont rejetés",
            r"die beschwerde ist abzuweisen, soweit darauf einzutreten ist",
            r"die beschwerde wird.*?abgewiesen",
            r"l'istanza di revisione.*?è respinta",
            r"die beschwerde wird im sinne der erwägungen abgewiesen",



            
        ],
    
    
          "berichtigt": [
                r"wird .*? berichtigt",
                r"werden .*? berichtigt",
                r"est rectifié",
                r"sont rectifiés",
            ],
           
        }     


 
    decision_textbox = soup.get_text("\n", strip=True)


    

    text = decision_textbox.lower()

    
    if (
        any(formel in text for formel in french)
        or re.search(
            r"^\s*ordonne\s*:\s*$",
            text,
            flags=re.MULTILINE
        )
    ):
        franc = franc + 1

    

    elif any(formel in text for formel in german):
        deutsch = deutsch + 1

    elif any(formel in text for formel in italy):
        ital = ital + 1
        
    



    teile = re.split(
        muster,
        text,
        maxsplit=1,
        flags=re.IGNORECASE | re.MULTILINE
    )
    

    
        #print(datei_name)
        #print(schluss_text)

    
    
    
    
    if len(teile) == 2 and teile[1].strip():
        schluss_text = teile[1].strip()


        korrekturen = {
            "le recours des rejeté": "le recours est rejeté",
            "die beschwerden werden gutheissen": "die beschwerden werden gutgeheissen",
             "le demande": "la demande",
            "le recours et irrecevable": "le recours est irrecevable",
            "les recours est irrecevable": "le recours est irrecevable",
            "die beschwerde teilweise gutgeheissen":  "die beschwerde wird teilweise gutgeheissen",

            "ammissibile, ricorso è respinto": "ammissibile, il ricorso è respinto",
            "la recours est rejeté": "le recours est rejeté",
            "w ird": "wird",
        }
        
        for falsch, richtig in korrekturen.items():
            schluss_text = schluss_text.replace(falsch, richtig)


    
    
        entscheid = entscheid_kategorisieren(schluss_text)
    
        
        if entscheid=="andere":
            unclear=unclear+1
           
            print(schluss_text)
            print(" ")
            print(".................................................................................................")
        
    
    else:
        print("Kein Text nach der Schlussformel:")
    




        
    tot = ital + franc + deutsch

 

    return  entscheid  

