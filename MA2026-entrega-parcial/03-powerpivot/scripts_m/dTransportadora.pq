// dTransportadora — dimensão de transportadoras (6 linhas)
let
    Fonte = cad_Etapa2{[Item = "dTransportadora", Kind = "Table"]}[Data],
    Tipos = Table.TransformColumnTypes(Fonte, {
        {"Transportadora", type text}, {"Tipo", type text}})
in
    Tipos
