Program som leser ATC kodetabeller i .xls-format. 

Bibliotektet som snakker med excel (xlrd) fungerer kun på .xls-filer. Dersom kodetabellen er i det nyere .xlsx-formatet må kodetabellen lagres på nytt i gammelt format.

atsim_del1: Bruker må angi sti til mappe med kodetabeller som skal leses inn.

atsim_del2: Kjøres etter at output-fila fra del 1 er oppdatert med segmenter. Eventuelt kan del 2 kjøres med en gang, men jobben med å definere segmenter blir litt vanskeligere.

atsim_func: Diverse funksjoner som brukes ellers i programmet.

atsim_class: Modulen inneholder klasser som kan ta en mengde kodetabeller og hente ut informasjonen. 

    -   Baliseoversikt(): "Permen" med alle kodetabellene du er interessert i. Innholder en liste over alle kodetabellene
    
    -   Kodetabell(): Hvert enkelt regneark, inneholder en liste over alle balisegruppene på arket
    
    -   Balisegruppe(): Den enkelte bgruppe, inneholder en liste over alle balisene i gruppa
    
    -   Balise(): En enkelt balise
    
    -   PD_table: En klasse for å printe ufullstendig informasjon i konsoll. Hovedsaklig for debugging.
