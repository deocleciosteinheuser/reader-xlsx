# reader-xlsx

Projeto para leitura de arquivo xlsx em PHP.

## Estrutura
- `src/`: Código-fonte PHP
- `public/`: Arquivos públicos (ex: index.php)

## Como rodar
Abra o terminal na raiz do projeto e execute:

```
php -d memory_limit=2048M -S localhost:8000 -t public
```

Acesse http://localhost:8000 no navegador para ver o "Hello World!".

## comando
```
(new XlsxToJson(new XlsxReaderBasic($file)))->toJson(
    [
        "grupo" => ["Código do Grupo","Nome do Grupo"],
        "classe" => ["Código da Classe", "Nome da Classe"],
        "subclasse" => ["Código do PDM", "Nome do PDM"]
    ]), 
    true
);
```

## dados
```
Código do Grupo	Nome do Grupo	Código da Classe	Nome da Classe	Código do PDM	Nome do PDM	Código do Item	Descrição do Item	Código NCM	
10	ARMAMENTO	1005	ARMAS DE FOGO DE CALIBRE ATÉ 120MM	1712	PEÇAS / ACESSÓRIOS ARMAMENTO	446820	PEÇAS / ACESSÓRIOS ARMAMENTO, MATERIAL:AÇO, TIPO:EIXO DA ALAVANCA DE MANEJO, REFERÊNCIA FABRIL:1015-15-020-1588	-	
```
## resposta
```
{
  "grupo": [
    {
      "Código do Grupo": "10",
      "Nome do Grupo": "ARMAMENTO",
      "classe": [
        {
          "Código da Classe": "1005",
          "Nome da Classe": "ARMAS DE FOGO DE CALIBRE ATÉ 120MM",
          "subclasse": [
            {
              "Código do PDM": "1712",
              "Nome do PDM": "PEÇAS / ACESSÓRIOS ARMAMENTO",
              "item": {
                "Código do Item": "627118",
                "Descrição do Item": "PEÇAS / ACESSÓRIOS ARMAMENTO, TIPO 5:Bloco da Culatra, MATERIAL:AÇO, APLICAÇÃO:MORTEIRO 120, REFERÊNCIA FABRIL 4:Spe-A1-0003",
                "Código NCM": "-"
              }
            }
          ]
        }
      ]
    }
  ]
}
```