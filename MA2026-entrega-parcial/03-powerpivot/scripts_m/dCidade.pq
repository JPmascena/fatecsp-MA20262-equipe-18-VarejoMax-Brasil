// dCidade — dimensão de cidades (24 cidades)
let
    Fonte = cad_Etapa2{[Item = "dCidade", Kind = "Table"]}[Data],
    Tipos = Table.TransformColumnTypes(Fonte, {
        {"Cidade", type text}, {"UF", type text}, {"Regiao", type text},
        {"Latitude", type number}, {"Longitude", type number}, {"Tem_Loja_Fisica", type text}})
in
    Tipos
