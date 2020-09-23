Az KS_Projekt_status_export inicializálásához a következő lépéseket kell végrehajtani:

1. Minden main mappában lévő .ps1 kiterjesztésű fájlt engedélyezni kell. Ennek módja: Jobb klikk/Tulajdonságok/Tiltás feloldása/OK (Ha nincs ilyen gomb akkor ez a lépés automatikusan megtörtént)
2. Futtatni az initializer.ps1 fájlt (operációs rendszer függvényében meg kell nézni, hogy szükség van-e a fájl eleján kikommentelt részekre)
3.0 Ha nem létezik akkor a main mappán belül létre kell hozni egy creds nevű mappát és azon belül egy "tavmeres-cred.txt" nevű fájlt.
3.1 Futtatni a pass_change.ps1 fájlt és a megkérdezett fiók esetében be kell írni az aktuálisan érvényes jelszót. (U.I.:Ha a jövőben változnak ezen fiókok jelszavai, akkor újra futtatni kell ezt a fájlt és be kell írni az aktuális jelszót.)
4. Le kell tölteni a node-ot (https://nodejs.org/en/download/), majd installálni. Installálás sikerességét a (lehetőleg rendszergazdai) konzolba beírt node -v paranccsal lehet ellenőrizni.
5. Ezek után a projekt mappájában (..../node-test) futtatni kell az npm install parancsot.
6. Ha minden jól ment a program készen áll a futtatásra (futtató kód konzolba: node index.js)