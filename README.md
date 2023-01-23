# autoresursy
Jest to program stworzony do łatwiejszego uzupełniania dokumentów dotyczących pracy mojego ojca.
Działa na podstawie wzorów w plikach .xlsx które niestety muszą być razem z programem w jednym folderze, nie potrafie tego inaczej rozwiązać.

Resursy to inaczej mówiąc okres zdolności użytkowej. Poprzez wpisanie danych wyliczane jest jak bardzo dany pojazd jest zużyty, czy wymaga przeglądu technicznego i czy jest bezpieczny. Protokół to wykaz wykonanych czynności i zużytych części podczas przeglądu/naprawy pojazdu.

Do działanie programu potrzebne są następujące biblioteki: Pillow (bez tego nie zapiszą sie obrazy), Openpyxl (edytuje i zapisuje nowe pliki xlsx ze wzorów) 
i pywin32 (zamienia pliki xlsx na pdf).

Przy uruchomieniu powinien stworzyć folder na pulpicie w którym będa sie zapisywać wszystkie dokumenty stworzone w programie.

Następnie pojawi się wybór jakie rozszerzenie ma mieć końcowy plik. Po wybraniu rozszerzenia nastąpi wybór typu resursu lub protokołu.
Po wybraniu program będzie pytał o informacje potrzebne do wypełnienia dokumentu. Nie trzeba martwic się wielkościa liter bo program je ignoruje i wpisuje do dokumentów zgodnie z normą.

W przypadku gdy wpisze sie prawdziwe dane program powinien np: wyznaczyc rok produkcji pojazdu na podstawie nr seryjnego(aby to zobaczyc nalezy w resursie wózka wpisać: Producent: linde, Numer seryjny: h2x393t) lub wyznaczyc udżwig wózka z typu (tu nalezy wpisac: Producent: linde, Typ: h25t).

