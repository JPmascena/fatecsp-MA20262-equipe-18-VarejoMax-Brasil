// dCliente — dimensão de clientes (1.607 clientes)
let
    Fonte = cad_Etapa2{[Item = "dCliente", Kind = "Table"]}[Data],
    Tipos = Table.TransformColumnTypes(Fonte, {
        {"ID_Cliente", type text}, {"Nome_Cliente", type text}, {"Email", type text},
        {"Data_Cadastro", type date}, {"Cidade", type text}, {"UF", type text}, {"Regiao", type text}})
in
    Tipos
