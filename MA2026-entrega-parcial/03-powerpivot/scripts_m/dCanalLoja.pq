// dCanalLoja — dimensão de canal × loja (3 online + 7 lojas físicas)
let
    Fonte = cad_Etapa2{[Item = "dCanalLoja", Kind = "Table"]}[Data],
    Tipos = Table.TransformColumnTypes(Fonte, {
        {"Chave_Canal_Loja", type text}, {"Canal", type text}, {"Loja", type text},
        {"Tipo_Canal", type text}, {"Cod_Loja", type text}, {"Cidade_Loja", type text}, {"UF_Loja", type text}})
in
    Tipos
