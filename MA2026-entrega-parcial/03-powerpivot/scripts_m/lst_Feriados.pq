// lst_Feriados — feriados nacionais do período (PROVISÓRIO: será substituído pela BrasilAPI)
let
    Fonte = #table(type table [Data = date, Feriado = text], {
        {#date(2025, 1, 1), "Confraternização Universal"},
        {#date(2025, 4, 21), "Tiradentes"},
        {#date(2025, 5, 1), "Dia do Trabalho"},
        {#date(2025, 9, 7), "Independência"},
        {#date(2025, 10, 12), "N. Sra. Aparecida"},
        {#date(2025, 11, 2), "Finados"},
        {#date(2025, 11, 15), "Proclamação da República"},
        {#date(2025, 12, 25), "Natal"},
        {#date(2026, 1, 1), "Confraternização Universal"},
        {#date(2026, 4, 21), "Tiradentes"},
        {#date(2026, 5, 1), "Dia do Trabalho"}
    })
in
    Fonte
