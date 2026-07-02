# Matematyka I stopnia

## Opracowanie wybranych odpowiedzi do pytań egzaminacyjnych

Wersja po polsku, ułożona według numerów wybranych pytań z prezentacji z repozytorium.

2 lipca 2026

## Jak czytać ten plik

Każdy blok zaczyna się od jednoznacznego nagłówka: **Pytanie N**. Numer jest numerem pytania z prezentacji, a nie numerem porządkowym w tym opracowaniu. Dlatego po pytaniu 1 od razu pojawia się pytanie 3, bo pytanie 2 zostało pominięte zgodnie z poleceniem.

W każdym pytaniu warto umieć powiedzieć: definicję lub główną ideę, najważniejsze własności, jeden typowy przykład oraz krótkie zdanie wyjaśniające sens tematu.

Przy opracowaniu wykorzystano korpus prezentacji z plików PDF znajdujących się w repozytorium. Materiał został scalony, ujednolicony i skrócony do formy odpowiedzi ustnej: bez slajdowej powtarzalności, ale z zachowaniem definicji, wzorów, przykładów i typowych miejsc, w których łatwo stracić punkty.

## Pytanie 1

**Brzmienie z listy:** *Pojęcie makra w Excelu, tworzenie, używanie i modyfikowanie makr.*

Makro w Excelu to sekwencja instrukcji automatyzujących określone czynności, dzięki czemu praca staje się szybsza, wygodniejsza i mniej podatna na błędy. W prezentacji podano dwa podstawowe sposoby tworzenia makr: przez **rejestrator makr** oraz przez programowanie w **VBA**. Rejestrator sprawdza się przy prostych, powtarzalnych zadaniach: zapisuje czynności użytkownika i generuje kod VBA, który później można edytować. Przy bardziej złożonych zadaniach korzysta się bezpośrednio z języka VBA. Skoroszyt z makrami powinien być zapisany w formacie **.XLSM**.

Język **VBA (Visual Basic for Applications)** służy do automatyzacji działań, tworzenia własnych poleceń i projektowania nowych funkcji arkusza. Wyróżnia się dwa podstawowe typy makr: **Sub** oraz **Function**. Procedura `Sub` działa jak nowe polecenie wykonywane przez użytkownika lub inne makro, natomiast `Function` zwraca wartość i może być używana zarówno w innych procedurach VBA, jak i bezpośrednio w formule arkusza. VBA operuje na obiektach ułożonych hierarchicznie, takich jak `Application`, `Workbook`, `Worksheet`, `Range` czy `Chart`, i korzysta z metod obiektów, zmiennych oraz konstrukcji sterujących, np. `If...Then`, `For...Next`, `With...End With`, `Select Case`.

Używanie i modyfikowanie makr obejmuje także kontakt z użytkownikiem i reagowanie na zdarzenia. W prezentacji omówiono okna dialogowe `InputBox`, `MsgBox` i formularze `UserForm`, a także kontrolki formularza i kontrolki `ActiveX`, którym można przypisywać makra. Excel pozwala również uruchamiać kod po wystąpieniu zdarzeń na poziomie skoroszytu, arkusza albo zdarzeń niepowiązanych bezpośrednio z obiektami. Dzięki temu makra służą nie tylko do pojedynczego „kliknięcia”, ale też do automatycznego sterowania zachowaniem pliku.

**Najważniejsze fakty.**

- Makro to sekwencja instrukcji automatyzujących pracę w Excelu.
- Dwa podstawowe sposoby tworzenia makr to **rejestrator makr** i **VBA**.
- Plik z makrami powinien być zapisany w formacie **.XLSM**.
- Rejestrator nadaje się do prostych, powtarzalnych czynności i generuje kod VBA.
- W VBA podstawowe typy procedur to **Sub** i **Function**.
- Najważniejsze obiekty VBA to `Application`, `Workbook`, `Worksheet`, `Range`, `Chart`.
- Do komunikacji z użytkownikiem służą m.in. `InputBox`, `MsgBox`, `UserForm`.
- Makra można wiązać z kontrolkami i zdarzeniami programu Excel.

**Przykład.** Użytkownik chce zautomatyzować kopiowanie zakresu komórek. Włącza **Deweloper -> Zarejestruj makro**, wykonuje kopiowanie, a następnie wybiera **Zatrzymaj rejestrowanie**. Excel tworzy kod VBA, który można później poprawić i przypisać np. do przycisku polecenia.

**Jak odpowiadać.** Najwygodniej iść w porządku: definicja makra -> dwa sposoby tworzenia -> rejestrator a VBA -> różnica między `Sub` i `Function` -> używanie przez okna dialogowe, kontrolki i zdarzenia -> typowe zastosowania.

**Uwaga na ustnym.** Nie utożsamiaj makra wyłącznie z rejestratorem. Trzeba podkreślić, że rejestrator tylko generuje kod VBA i nie nadaje się do bardziej złożonych algorytmów. Warto też pamiętać o różnicy `Sub` / `Function` oraz o formacie **.XLSM**.

## Pytanie 3

**Brzmienie z listy:** *Metoda pajęczyny, opis i przykład stosowania.*

Metoda pajęczyny jest graficzną metodą analizy iteracji w dyskretnym układzie dynamicznym opisanym równaniem różnicowym

$$
x_{n+1} = f(x_n).
$$

Badamy tu zachowanie ciągu $(x_n)$: czy jest zbieżny, czy oscyluje, czy się rozbiega. Podstawowym pojęciem jest **punkt równowagi** (punkt stały) $x^\*$, spełniający warunek

$$
f(x^\*) = x^\*.
$$

Geometrycznie jest to punkt przecięcia wykresów $y = f(x)$ oraz $y = x$.

Konstrukcja metody jest następująca: startujemy od punktu początkowego $x_0$, prowadzimy pionowo do wykresu $y = f(x)$ i otrzymujemy punkt $(x_0, x_1)$, gdzie $x_1 = f(x_0)$. Następnie prowadzimy poziomo do prostej $y = x$, czyli do punktu $(x_1, x_1)$, potem znowu pionowo do wykresu $y = f(x)$, uzyskując $(x_1, x_2)$, i tak dalej. W ten sposób powstaje „pajęczyna”, która pokazuje kolejne iteracje i pozwala wizualnie ocenić, czy trajektoria zbliża się do punktu równowagi, czy od niego ucieka.

Jeżeli funkcja $f$ jest różniczkowalna w punkcie równowagi, to lokalne kryterium stabilności ma postać:

$$
|f'(x^\*)| < 1 \Rightarrow \text{punkt stabilny}, \qquad
|f'(x^\*)| > 1 \Rightarrow \text{punkt niestabilny}.
$$

Metoda pajęczyny pozwala to zobaczyć geometrycznie: jeśli kolejne kroki „zawijają się” do punktu przecięcia, iteracja jest stabilna; jeśli się oddalają, punkt jest niestabilny.

**Najważniejsze fakty.**

- Rozważamy model dyskretny postaci $x_{n+1} = f(x_n)$.
- Jest to szczególny przypadek równania różnicowego.
- Punkt równowagi spełnia warunek $f(x^\*) = x^\*$.
- Graficznie punkt równowagi to przecięcie wykresów $y = f(x)$ i $y = x$.
- Metoda pajęczyny polega na naprzemiennym rysowaniu odcinków pionowych i poziomych.
- Służy do badania zbieżności, stabilności i charakteru iteracji.
- Kryterium lokalne: $|f'(x^\*)| < 1$ daje stabilność, a $|f'(x^\*)| > 1$ niestabilność.

**Przykład.** Rysujemy wykresy $y = f(x)$ i $y = x$. Dla zadanego $x_0$ konstruujemy kolejno punkty $(x_0, x_1)$, $(x_1, x_1)$, $(x_1, x_2)$, $(x_2, x_2)$ itd. Jeżeli pajęczyna zawija się do punktu przecięcia obu wykresów, to ciąg $(x_n)$ dąży do stabilnego punktu równowagi.

**Jak odpowiadać.** Najprostszy schemat to: model dyskretny -> równanie $x_{n+1} = f(x_n)$ -> definicja punktu równowagi -> opis konstrukcji krok po kroku -> warunek stabilności z pochodną.

**Uwaga na ustnym.** Najczęstszy błąd to pominięcie prostej $y = x$. W tej metodzie nie wystarczy patrzeć na sam wykres $y = f(x)$; punkt równowagi wyznacza się właśnie jako przecięcie z $y = x$. Trzeba też pamiętać o module w warunku $|f'(x^\*)| < 1$.

## Pytanie 4

**Brzmienie z listy:** *Liczba losowa a pseudolosowa. Sposoby generowania liczb losowych i pseudolosowych.*

Ciąg liczb losowych to taki ciąg, który jest **równomiernie rozłożony** na wszystkie możliwe wartości, a jego kolejne elementy są **niezależne** od wcześniej wygenerowanych. Generatory liczb prawdziwie losowych korzystają ze źródeł entropii, czyli zjawisk rzeczywistych o dużej nieprzewidywalności, takich jak rzut monetą, losowanie kul z urny, ruletka, szum termiczny, dane atmosferyczne czy rozpad promieniotwórczy.

Liczby pseudolosowe są natomiast imitacją losowości. Powstają algorytmicznie, bez odwołania do zjawisk fizycznych. Typowy generator pseudolosowy opisuje się przez strukturę $(S,\mu,f,U,g)$: $S$ jest skończonym zbiorem stanów, $\mu$ rozkładem początkowym, $f$ funkcją przejścia, $U$ zbiorem wyników, a $g$ funkcją wynikową. Generator startuje od **ziarna** (*seed*), a ponieważ przestrzeń stanów jest skończona, po pewnym czasie ciąg zaczyna się powtarzać; stąd pojęcie **okresu** generatora.

Ważną klasą są **generatory kongruencyjne**, wykorzystujące operację modulo. Najprostszy liniowy generator mieszany ma postać

$$
X_n = (aX_{n-1} + c) \bmod M,
$$

a jego szczególny przypadek stanowi generator multiplikatywny

$$
X_n = aX_{n-1} \bmod M.
$$

W prezentacji podano też warunki pełnego okresu dla generatora mieszanego: moduł $M$ i stała $c$ muszą być względnie pierwsze, każdy czynnik pierwszy liczby $M$ musi dzielić $a-1$, a jeśli $4 \mid M$, to także $4 \mid (a-1)$. Omawiane są również uogólnienia: generator liniowy uogólniony, Fibonacciego, generator kwadratowy i inwersyjny.

Z liczb jednostajnych generuje się dalej liczby o innych rozkładach. Pierwsza metoda to **metoda odwrotnej dystrybuanty**: dla $U \sim U(0,1)$ i ściśle rosnącej dystrybuanty $F$ przyjmujemy $X = F^{-1}(U)$. Druga to **metoda eliminacji von Neumanna**: losujemy parę $(x,U)$ i akceptujemy $x$ tylko wtedy, gdy zachodzi warunek $U < f(x)$. Metoda odwrotnej dystrybuanty wymaga znajomości $F$ i $F^{-1}$, natomiast metoda von Neumanna działa także wtedy, gdy nie znamy dystrybuanty, ale w wersji podstawowej bywa mało wydajna.

**Najważniejsze fakty.**

- Prawdziwe liczby losowe pochodzą ze źródeł entropii, czyli zjawisk fizycznych lub mechanicznych.
- Liczby pseudolosowe są generowane deterministycznie przez algorytm.
- Generator pseudolosowy startuje od **ziarna** i ma skończony **okres**.
- Typowa struktura generatora ma postać $(S,\mu,f,U,g)$.
- Generatory kongruencyjne używają operacji **modulo**.
- Liniowy generator mieszany ma wzór $X_n = (aX_{n-1} + c) \bmod M$.
- Dla pełnego okresu generatora mieszanego trzeba spełnić warunki na $M$, $c$ i $a-1$.
- Inne rozkłady można generować metodą odwrotnej dystrybuanty albo metodą eliminacji von Neumanna.

**Przykład.** Dla generatora liniowego z parametrami

$$
M = 8,\quad a = 5,\quad c = 3,\quad X_0 = 7
$$

otrzymujemy kolejno

$$
X_1 = 6,\; X_2 = 1,\; X_3 = 0,\; X_4 = 3,\; X_5 = 2,\; X_6 = 5,\; X_7 = 4,\; X_8 = 7.
$$

Po ośmiu krokach wracamy do wartości początkowej, więc generator ma maksymalny okres $p = 8 = M$.

**Jak odpowiadać.** Najlepiej zacząć od wyraźnego rozróżnienia: losowe = zjawiska rzeczywiste, pseudolosowe = algorytm. Potem warto podać pojęcia ziarna i okresu, wzór generatora kongruencyjnego, a na końcu wspomnieć o metodzie odwrotnej dystrybuanty i metodzie von Neumanna.

**Uwaga na ustnym.** Nie wolno mówić, że liczby pseudolosowe są losowe w sensie ścisłym: po ustaleniu ziarna generator działa deterministycznie. Dobrze też pamiętać, że metoda odwrotnej dystrybuanty wymaga znajomości $F$ i $F^{-1}$, a metoda von Neumanna może być kosztowna obliczeniowo.

## Pytanie 5

**Brzmienie z listy:** *Łańcuch Markowa na czasie dyskretnym i przestrzeni dyskretnej.*

Proces stochastyczny $\{X_n\}_{n \in \mathbb{N}_0}$ o czasie dyskretnym i przeliczalnej przestrzeni stanów $Z$ nazywamy łańcuchem Markowa, jeżeli spełnia własność Markowa:

$$
P(X_{n+1} = j \mid X_n = i, X_{n-1} = i_{n-1}, \dots, X_0 = i_0)
=
P(X_{n+1} = j \mid X_n = i).
$$

Znaczy to, że przyszłość zależy tylko od stanu aktualnego, a nie od całej historii procesu. Jeżeli prawdopodobieństwa przejścia nie zależą od chwili $n$, to łańcuch jest **jednorodny w czasie**.

Przy skończonej przestrzeni stanów model opisują dwa podstawowe elementy: rozkład początkowy oraz macierz przejścia $P = [p_{ij}]$, gdzie

$$
p_{ij} = P(X_{n+1} = j \mid X_n = i).
$$

Macierz ta jest **stochastyczna**, czyli ma elementy nieujemne i sumy wierszy równe $1$. Prawdopodobieństwa przejścia w $n$ krokach tworzą macierz

$$
P^{(n)} = P^n,
$$

a równania Chapmana-Kołmogorowa mają postać

$$
p_{ij}^{(n+m)} = \sum_{k \in Z} p_{ik}^{(m)} p_{kj}^{(n)}.
$$

Dalsza analiza dotyczy struktury przestrzeni stanów. Piszemy $i \to j$, gdy ze stanu $i$ można dojść do $j$ z dodatnim prawdopodobieństwem po skończonej liczbie kroków, oraz $i \leftrightarrow j$, gdy osiągalność jest wzajemna. Relacja komunikowania dzieli przestrzeń stanów na **klasy komunikujące się**. Szczególne znaczenie mają klasy zamknięte, stany pochłaniające oraz **czas pierwszego osiągnięcia** zbioru

$$
T_A = \inf\{n \ge 0 : X_n \in A\},
$$

który może być skończony albo równy $\infty$.

**Najważniejsze fakty.**

- Czas jest dyskretny: $n = 0,1,2,\dots$.
- Przestrzeń stanów jest dyskretna, czyli co najwyżej przeliczalna.
- Własność Markowa oznacza, że przyszłość zależy tylko od stanu bieżącego.
- Łańcuch jednorodny spełnia $p_{ij}(n) = p_{ij}$.
- Macierz przejścia jest stochastyczna: $p_{ij} \ge 0$ oraz $\sum_j p_{ij} = 1$.
- Macierz przejść w $n$ krokach to $P^{(n)} = P^n$.
- Relacja komunikowania dzieli stany na klasy komunikujące się.
- Klasa zamknięta nie może zostać opuszczona.
- Stan pochłaniający spełnia warunek $p_{ii} = 1$.
- Łańcuch jest nieredukowalny, gdy każdy stan komunikuje się z każdym innym.

**Przykład.** Rozważmy łańcuch o stanach $\{1,2,3\}$, w którym stany $1$ i $3$ są pochłaniające, a ze stanu $2$ przechodzimy do $1$ z prawdopodobieństwem $0{,}7$ i do $3$ z prawdopodobieństwem $0{,}3$. Dla celu $A = \{1\}$ i startu z $2$ mamy

$$
P(T_A = \infty \mid X_0 = 2) = 0{,}3,
$$

bo z prawdopodobieństwem $0{,}3$ proces wpada od razu do stanu $3$, z którego nie da się już dojść do $1$.

**Jak odpowiadać.** Najczytelniej iść w kolejności: definicja procesu -> własność Markowa -> jednorodność -> macierz przejścia -> przejścia wielokrokowe -> klasy stanów -> czas osiągnięcia zbioru.

**Uwaga na ustnym.** Nie mieszaj prawdopodobieństw jednokrokowych $p_{ij}$ z $n$-krokowymi $p_{ij}^{(n)}$. Warto też pamiętać, że stan pochłaniający jest szczególnym przypadkiem klasy zamkniętej, a czas osiągnięcia zbioru nie musi być prawie na pewno skończony.

## Pytanie 6

**Brzmienie z listy:** *Pojęcie stabilności stanów stacjonarnych równania różniczkowego.*

Punktem wyjścia jest **stabilność w sensie Lapunowa**. Rozwiązanie $x(t)$ układu

$$
x'(t) = f(t,x(t))
$$

nazywamy stabilnym, jeżeli każde rozwiązanie zaczynające dostatecznie blisko $x(t)$ pozostaje potem dowolnie blisko niego. Formalnie: dla każdego $\varepsilon > 0$ można dobrać takie $\delta > 0$, aby małe zaburzenie warunku początkowego nie wyprowadzało rozwiązania poza $\varepsilon$-otoczenie rozwiązania badanego. Jeżeli dodatkowo

$$
\lim_{t \to \infty} \|x(t) - \widetilde{x}(t)\| = 0,
$$

to rozwiązanie jest **asymptotycznie stabilne**. Jeśli nie jest stabilne, mówimy o **niestabilności**.

W przypadku równań autonomicznych

$$
x'(t) = f(x(t))
$$

pojęcie jest szczególnie naturalne, bo rozwiązania stałe odpowiadają punktom równowagi. Dla układu planarnego

$$
\begin{cases}
x'(t) = f(x,y),\\
y'(t) = g(x,y),
\end{cases}
$$

punkt $(x^\*, y^\*)$, dla którego $f(x^\*, y^\*) = 0$ i $g(x^\*, y^\*) = 0$, nazywamy **punktem krytycznym**, **punktem równowagi** albo **rozwiązaniem stacjonarnym**. Odpowiada mu rozwiązanie stałe. Prezentacja zwraca też uwagę, że równanie nieautonomiczne można sprowadzić do autonomicznego przez dołączenie zmiennej czasu jako kolejnej składowej układu.

Dla układów liniowych

$$
y'(t) = Ay(t)
$$

klasyfikacja punktu równowagi opiera się na wartościach własnych macierzy $A$. Dwie rzeczywiste wartości własne ujemne dają węzeł asymptotycznie stabilny, dwie dodatnie węzeł niestabilny, a wartości o przeciwnych znakach siodło, więc niestabilność. Jeżeli $\lambda_{1,2} = \alpha \pm \beta i$, to znak części rzeczywistej rozstrzyga o zachowaniu: $\alpha < 0$ daje ognisko asymptotycznie stabilne, $\alpha > 0$ ognisko niestabilne, a $\alpha = 0$ centrum stabilne, ale nie asymptotycznie stabilne.

**Najważniejsze fakty.**

- Stabilność w sensie Lapunowa oznacza trwałość względem małych zaburzeń.
- Asymptotyczna stabilność = stabilność + zbieżność trajektorii do rozwiązania badanego.
- Niestabilność jest zaprzeczeniem stabilności.
- W równaniu autonomicznym prawa strona nie zależy jawnie od czasu.
- Punkt stacjonarny to miejsce zerowe prawej strony układu.
- Punkt stacjonarny odpowiada rozwiązaniu stałemu.
- Równanie nieautonomiczne można sprowadzić do autonomicznego przez dołączenie zmiennej czasu.
- W układzie liniowym o stabilności decydują wartości własne macierzy.
- Ujemne części rzeczywiste wartości własnych oznaczają stabilność asymptotyczną.
- Centrum jest stabilne, ale nie asymptotycznie stabilne.

**Przykład.** Dla układu

$$
y'(t) = Ay(t), \qquad
A =
\begin{bmatrix}
-1 & 0 \\
0 & -2
\end{bmatrix}
$$

wartości własne wynoszą $\lambda_1 = -1$, $\lambda_2 = -2$. Obie są ujemne, więc punkt równowagi $y = 0$ jest węzłem asymptotycznie stabilnym: trajektorie bliskie zera nie tylko pozostają blisko, ale z czasem do zera dążą.

**Jak odpowiadać.** Najlepiej zacząć od definicji stabilności Lapunowa, od razu odróżnić ją od stabilności asymptotycznej, następnie zdefiniować punkt stacjonarny jako rozwiązanie stałe i zakończyć klasyfikacją przez wartości własne.

**Uwaga na ustnym.** Nie utożsamiaj słów „stabilny” i „asymptotycznie stabilny” - to nie są synonimy. Najbezpieczniejsza reguła pamięciowa brzmi: dla układu liniowego stabilność asymptotyczną zapewnia ujemna część rzeczywista każdej wartości własnej.

## Pytanie 7

**Brzmienie z listy:** *Model SIR, schemat modelu.*

Model SIR jest klasycznym modelem rozprzestrzeniania się choroby zakaźnej w populacji. Populację dzieli się na trzy klasy:

- $S$ - osobnicy podatni na zakażenie (*susceptible*),
- $I$ - osobnicy zakażeni i zakażający (*infected*),
- $R$ - osobnicy wyzdrowiali i odporni albo usunięci z populacji (*resistant* / *removed*).

Schemat przejść ma postać

$$
S \xrightarrow{\alpha} I \xrightarrow{\gamma} R,
$$

gdzie $\alpha$ jest współczynnikiem rozprzestrzeniania się infekcji, a $\gamma$ współczynnikiem zdrowienia.

W wersji przedstawionej w prezentacji zakłada się stałą całkowitą liczebność populacji,

$$
S(t) + I(t) + R(t) = N,
$$

brak rozrodczości, brak okresu inkubacji oraz jednakowe prawdopodobieństwo kontaktu każdej pary osobników. Odpowiada temu układ

$$
\frac{dS}{dt} = -\alpha SI, \qquad
\frac{dI}{dt} = \alpha SI - \gamma I, \qquad
\frac{dR}{dt} = \gamma I,
$$

z warunkami początkowymi $S(0) = S_0 > 0$, $I(0) = I_0 > 0$, $R(0) = 0$. Z pierwszego równania wynika, że $S(t)$ jest funkcją nierosnącą.

Kluczowe jest równanie dla klasy zakażonych:

$$
I'(t) = I(t)(\alpha S(t) - \gamma).
$$

Stąd początkowy rozwój epidemii zależy od porównania $S_0$ z progiem $\gamma / \alpha$. Jeżeli $S_0 \le \gamma / \alpha$, liczba zakażonych maleje od początku. Jeżeli $S_0 > \gamma / \alpha$, liczba zakażonych początkowo rośnie, aż liczba podatnych spadnie do poziomu $\gamma / \alpha$.

W prezentacji próg epidemiczny opisuje podstawowy współczynnik reprodukcji

$$
R_0 = \frac{\alpha}{\gamma} S(0).
$$

Jeżeli $R_0 < 1$, choroba zanika, a jeżeli $R_0 > 1$, dochodzi do wybuchu epidemii.

**Najważniejsze fakty.**

- Klasy modelu to $S$ - podatni, $I$ - zakażeni, $R$ - odporni/usunięci.
- Parametr $\alpha$ opisuje transmisję, a $\gamma$ zdrowienie.
- Populacja jest stała: $S + I + R = N$.
- Zakażenia powstają z intensywnością proporcjonalną do $\alpha SI$.
- Wyzdrowienia zachodzą z intensywnością proporcjonalną do $\gamma I$.
- W modelu nie ma okresu inkubacji.
- Zakłada się jednorodne mieszanie populacji.
- Funkcja $S(t)$ jest nierosnąca.
- Liczba zakażonych rośnie wtedy, gdy $\alpha S(t) - \gamma > 0$.
- W tej wersji modelu prezentacja zapisuje $R_0$ jako $\frac{\alpha}{\gamma}S(0)$.

**Przykład.** Jeśli na początku

$$
S_0 = \frac{2\gamma}{\alpha},
$$

to

$$
R_0 = \frac{\alpha}{\gamma} S_0 = 2 > 1.
$$

W takiej sytuacji mamy $I'(0) > 0$, więc liczba zakażonych początkowo rośnie. Wzrost trwa do chwili, gdy liczba podatnych spadnie do poziomu $\gamma / \alpha$.

**Jak odpowiadać.** Najlepiej mówić w kolejności: znaczenie klas $S,I,R$ -> założenia modelu -> układ równań -> interpretacja znaków w równaniach -> próg epidemiczny i $R_0$.

**Uwaga na ustnym.** Nie mieszaj tej wersji z modelami z demografią albo okresem inkubacji. Warto też zaznaczyć, że w różnych normalizacjach modelu wzór na $R_0$ może wyglądać inaczej; tutaj pracujemy na liczebnościach klas.

## Pytanie 8

**Brzmienie z listy:** *Modele epidemiologiczne, przykłady.*

Modele epidemiologiczne są matematycznym opisem rozwoju chorób zakaźnych w czasie. Ich celem jest zrozumienie dynamiki zakażeń, prognozowanie liczby zachorowań oraz ocena skuteczności interwencji, takich jak izolacja, dystans społeczny czy szczepienia. Podstawowa idea polega na podziale populacji na klasy odpowiadające stanowi zdrowia, a następnie opisaniu przejść między tymi klasami za pomocą równań różniczkowych.

Najprostsze modele to **SI**, **SIS** i **SIR**. W modelu SI osoby przechodzą tylko z klasy podatnych $S$ do zakażonych $I$. W SIS po wyzdrowieniu wracają do klasy podatnych, więc choroba nie daje trwałej odporności. W SIR pojawia się klasa $R$, obejmująca osoby wyzdrowiałe lub usunięte z populacji, co odpowiada nabyciu odporności albo zakończeniu udziału w transmisji. Modele **SEIR** i **SEIS** dodają jeszcze klasę $E$, czyli fazę utajoną: osoba jest już zakażona, ale jeszcze nie zakaża innych.

Przykładowe układy z prezentacji mają postać:

$$
\text{SI:}\quad
\begin{cases}
S' = -\alpha IS,\\
I' = \alpha IS,
\end{cases}
$$

$$
\text{SIS:}\quad
\begin{cases}
S' = -\alpha IS + \beta I,\\
I' = \alpha IS - \beta I,
\end{cases}
$$

$$
\text{SIR:}\quad
\begin{cases}
S' = -\alpha IS,\\
I' = \alpha IS - \gamma I,\\
R' = \gamma I.
\end{cases}
$$

W modelu SEIR dochodzi klasa $E$ z przejściem do zakaźności z tempem $\sigma$, a w bardziej rozbudowanych modelach można uwzględniać także rozrodczość, śmiertelność i utratę odporności, np. w modelu SIRS.

Kluczowym pojęciem jest **podstawowa liczba odtwarzania** $R_0$, czyli średnia liczba osób zakażonych przez jedną osobę chorą w populacji bez odporności. W prostym modelu SIR w prezentacji zapisano

$$
R_0 = \frac{\alpha}{\gamma},
$$

a ogólna zasada brzmi:

$$
R_0 = (\text{tempo zakażania}) \times (\text{średni czas zakaźności}).
$$

Jeżeli $R_0 < 1$, choroba zanika; jeżeli $R_0 > 1$, możliwy jest rozwój epidemii. W modelach bardziej złożonych $R_0$ wyznacza się np. metodą macierzy następnego pokolenia.

**Najważniejsze fakty.**

- Model epidemiologiczny dzieli populację na klasy stanów choroby i opisuje przejścia między nimi.
- **SI**: tylko $S \to I$.
- **SIS**: po wyzdrowieniu powrót do $S$, brak trwałej odporności.
- **SIR**: po zakażeniu przejście do $R$, czyli odporność albo usunięcie z populacji.
- **SEIR / SEIS**: dodatkowa klasa $E$ opisuje okres inkubacji.
- W prostym SIR podstawowa liczba odtwarzania ma postać $R_0 = \alpha / \gamma$.
- Ogólna zasada to $R_0 = (\text{tempo zakażania}) \times (\text{średni czas zakaźności})$.
- Dla $R_0 < 1$ epidemia wygasa, a dla $R_0 > 1$ może się rozwinąć.
- Modele można rozszerzać o narodziny, zgony i utratę odporności.

**Przykład.** W modelu SIR populację dzielimy na trzy grupy: $S$ - podatni, $I$ - zakażeni, $R$ - wyzdrowiali/usunięci. Układ

$$
S' = -\alpha IS,\qquad I' = \alpha IS - \gamma I,\qquad R' = \gamma I
$$

opisuje pełną dynamikę. Jeśli $\alpha / \gamma > 1$, liczba zakażonych może początkowo rosnąć; jeśli $\alpha / \gamma < 1$, epidemia nie utrzyma się w populacji.

**Jak odpowiadać.** Najlepiej zacząć od ogólnej definicji modelu epidemiologicznego i celu modelowania. Następnie przejść do klasyfikacji SI, SIS, SIR, SEIR, SEIS, a potem podać znaczenie $R_0$ i zasadę progową.

**Uwaga na ustnym.** Nie wolno mylić modeli SIS i SIR: w pierwszym nie ma trwałej odporności, w drugim jest. Trzeba też pamiętać, że wzór $R_0 = \alpha / \gamma$ dotyczy prostego modelu SIR, a nie wszystkich modeli epidemiologicznych.

## Pytanie 9

**Brzmienie z listy:** *Metody Eulera i Rungego-Kutty rozwiązywania równań różniczkowych zwyczajnych.*

Punktem wyjścia jest zagadnienie początkowe

$$
x'(t) = f(t,x(t)), \qquad x(t_0) = x_0.
$$

Metody numeryczne nie szukają wzoru dokładnego na rozwiązanie, tylko wyznaczają kolejne przybliżenia $x_i \approx x(t_i)$ w punktach siatki

$$
t_i = t_0 + ih,
$$

gdzie $h > 0$ jest krokiem całkowania. Najprostsza jest **jawna metoda Eulera**:

$$
x_{i+1} = x_i + h f(t_i,x_i).
$$

Interpretacja geometryczna jest prosta: na krótkim odcinku zastępujemy rozwiązanie styczną w lewym końcu kroku. Jest to metoda pierwszego rzędu: błąd lokalny ma rząd $O(h^2)$, a błąd globalny $O(h)$.

Metody **Rungego-Kutty** poprawiają dokładność, bo w jednym kroku wykorzystują nie jedno, lecz kilka nachyleń. W prezentacji omówiono metodę **Heuna** i **metodę punktu środkowego**, które są metodami RK rzędu 2, oraz klasyczną metodę **RK4**:

$$
F_1 = h f(t_i,x_i),
$$

$$
F_2 = h f\left(t_i + \frac{h}{2}, x_i + \frac{F_1}{2}\right),
$$

$$
F_3 = h f\left(t_i + \frac{h}{2}, x_i + \frac{F_2}{2}\right),
$$

$$
F_4 = h f(t_i + h, x_i + F_3),
$$

$$
x_{i+1} = x_i + \frac{1}{6}(F_1 + 2F_2 + 2F_3 + F_4).
$$

Klasyczna RK4 ma błąd lokalny rzędu $O(h^5)$ i błąd globalny rzędu $O(h^4)$. W praktyce oznacza to dużo lepszą dokładność niż Euler, ale kosztem większej liczby obliczeń funkcji $f$.

Ważna jest także stabilność. Dla równania testowego

$$
x' = \lambda x
$$

metoda Eulera daje

$$
x_n = (1 + h\lambda)^n
$$

i jest stabilna, gdy

$$
|1 + h\lambda| < 1.
$$

Jeżeli $\lambda < 0$, to warunek przyjmuje postać

$$
0 < h < -\frac{2}{\lambda}.
$$

Dlatego dla równań sztywnych Euler jawny może wymagać bardzo małego kroku.

**Najważniejsze fakty.**

- Zagadnienie początkowe ma postać $x'(t) = f(t,x(t))$, $x(t_0) = x_0$.
- Przybliżenia liczymy w punktach siatki $t_i = t_0 + ih$.
- Jawna metoda Eulera ma wzór $x_{i+1} = x_i + h f(t_i,x_i)$.
- Euler jest metodą pierwszego rzędu: błąd lokalny $O(h^2)$, globalny $O(h)$.
- Heun i metoda punktu środkowego są metodami RK2.
- RK2 ma błąd lokalny $O(h^3)$ i globalny $O(h^2)$.
- Klasyczna RK4 ma błąd lokalny $O(h^5)$ i globalny $O(h^4)$.
- Dla równania testowego $x' = \lambda x$ Euler jest stabilny, gdy $|1+h\lambda|<1$.
- Większa dokładność zwykle wymaga większej liczby obliczeń na krok.

**Przykład.** Dla równania

$$
x' = t + x, \qquad x(0) = 1, \qquad h = 0{,}1
$$

metoda Eulera daje po jednym kroku

$$
x_1 = 1 + 0{,}1 \cdot f(0,1) = 1 + 0{,}1 \cdot 1 = 1{,}1000.
$$

W tym samym zadaniu klasyczna metoda RK4 daje

$$
x_1 \approx 1{,}11034,
$$

czyli praktycznie wartość dokładną $x(0{,}1)$ z prezentacji.

**Jak odpowiadać.** Najpierw trzeba zdefiniować zagadnienie początkowe i ideę przejścia krok po kroku. Potem podać wzór Eulera, jego interpretację i rząd dokładności. Następnie wyjaśnić ideę Rungego-Kutty: kilka nachyleń zamiast jednego, po czym porównać RK2 i RK4.

**Uwaga na ustnym.** Najczęstszy błąd to mylenie błędu lokalnego z globalnym. Trzeba też pamiętać, że większy rząd metody nie daje lepszego wyniku „za darmo” - płaci się większą liczbą obliczeń funkcji $f$ w każdym kroku.

## Pytanie 10

**Brzmienie z listy:** *Metody wyznaczania zer nieliniowych funkcji skalarnych.*

Problem polega na rozwiązaniu równania

$$
f(x) = 0,
$$

czyli na znalezieniu takiej liczby $r$, że $f(r) = 0$. Dla funkcji nieliniowych rozwiązania dokładnego często nie da się otrzymać w postaci jawnej, dlatego stosuje się metody iteracyjne, które budują ciąg przybliżeń

$$
x_0, x_1, x_2, \dots \to r.
$$

W prezentacji omówiono trzy podstawowe metody: Newtona, siecznych i bisekcji.

**Metoda Newtona** zastępuje funkcję styczną w punkcie $(x_n,f(x_n))$, a jej punkt przecięcia z osią $OX$ przyjmuje jako nowe przybliżenie:

$$
x_{n+1} = x_n - \frac{f(x_n)}{f'(x_n)}.
$$

Wymaga więc jednego punktu startowego i znajomości pochodnej. Dla pierwiastka pojedynczego, przy funkcji klasy $C^2$ i dobrym punkcie początkowym, ma zbieżność kwadratową:

$$
|\varepsilon_{n+1}| \le C |\varepsilon_n|^2.
$$

**Metoda siecznych** nie potrzebuje pochodnej, bo zastępuje ją ilorazem różnicowym:

$$
x_{n+1} = x_n - \frac{f(x_n)(x_n - x_{n-1})}{f(x_n) - f(x_{n-1})}.
$$

Wymaga dwóch punktów startowych i ma zbieżność nadliniową rzędu

$$
\alpha = \frac{1+\sqrt{5}}{2} \approx 1{,}62.
$$

**Metoda bisekcji** jest najbezpieczniejsza. Zakłada ciągłość funkcji na przedziale $[a,b]$ oraz zmianę znaku:

$$
f(a)f(b) < 0.
$$

Wtedy w każdej iteracji bierzemy środek przedziału i wybieramy tę połowę, w której znak się zmienia. Metoda jest zawsze zbieżna przy spełnionych założeniach, ale tylko liniowo. Dla środka przedziału $c_n$ mamy oszacowanie

$$
|r - c_n| \le 2^{-(n+1)}(b_0 - a_0).
$$

Zatem: Newton jest najszybszy, sieczne są kompromisem, a bisekcja daje największą pewność zbieżności.

**Najważniejsze fakty.**

- Szukamy rozwiązania równania $f(x)=0$.
- Metoda Newtona ma wzór $x_{n+1} = x_n - \frac{f(x_n)}{f'(x_n)}$.
- Newton wymaga jednego punktu startowego i znajomości pochodnej.
- Dla pierwiastka pojedynczego Newton ma zbieżność kwadratową.
- Metoda siecznych nie wymaga pochodnej, ale potrzebuje dwóch punktów początkowych.
- Rząd zbieżności metody siecznych wynosi około $1{,}62$.
- Metoda bisekcji wymaga ciągłości funkcji i warunku $f(a)f(b)<0$.
- W bisekcji po każdym kroku długość przedziału maleje dwukrotnie.
- Bisekcja jest najwolniejsza, ale najbardziej niezawodna.

**Przykład.** Wyznaczamy $\sqrt{2}$ z równania

$$
x^2 - 2 = 0.
$$

Dla metody Newtona bierzemy $f(x) = x^2 - 2$ i $f'(x) = 2x$, więc

$$
x_{n+1} = \frac{1}{2}\left(x_n + \frac{2}{x_n}\right).
$$

Przy starcie $x_0 = 1$ dostajemy szybko:

$$
x_1 = 1{,}5,\quad
x_2 = 1{,}4166666667,\quad
x_3 = 1{,}4142156863,\quad
x_4 = 1{,}4142135624.
$$

To dobrze ilustruje kwadratową zbieżność Newtona.

**Jak odpowiadać.** Najlepiej zacząć od ogólnego problemu $f(x)=0$, potem omówić po kolei Newtona, sieczne i bisekcję: wzór, wymagania, szybkość zbieżności. Na końcu warto dodać krótkie porównanie praktyczne: szybkość kontra pewność zbieżności.

**Uwaga na ustnym.** Nie wolno pomijać założeń. Dla Newtona trzeba pamiętać o pochodnej i dobrym punkcie startowym, dla siecznych o dwóch punktach początkowych, a dla bisekcji o ciągłości funkcji i zmianie znaku na końcach przedziału.
