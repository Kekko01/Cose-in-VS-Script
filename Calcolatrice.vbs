Do

calcolo=inputbox ("Inserisci il calcolo. (Non inserire l'uguale alla fine)") 
msgbox eval(calcolo)
Loop Until trim(ucase(calcolo))= "BASTA"