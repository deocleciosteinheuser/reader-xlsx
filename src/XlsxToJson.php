<?php
namespace ReaderXlsx;

class XlsxToJson
{
    private $reader;


    /**
     * Construtor da classe XlsxToJson
     *
     * @param XlsxReaderBasic $reader Instância do leitor de XLSX
     */
    public function __construct(XlsxReaderBasic $reader)
    {
        $this->reader = $reader;
    }

    /**
     * Converte o conteúdo do XLSX em JSON.
     *
     * @param array $grupo Array de colunas para agrupamento hierárquico (opcional)
     * @return string JSON formatado dos dados da planilha
     */
    public function toJson(array $grupos = []): string
    {
        $headers = $this->reader->getHeaders();
        $result = [];
        $first = true;
        foreach ($this->reader->getRows() as $row) {
            if ($first) { $first = false; continue; } // pula header
            $assoc = [];
            foreach ($headers as $i => $header) {
                $assoc[$header] = $row[$i] ?: null;
            }
            if ($grupos) {
                $this->addToGroup($result, $assoc, $grupos);
            } else {
                $result[] = $assoc;
            }
        }
        // remove codigo dos grupos
        $this->arrayvalues($result, array_keys($grupos));
        // Retorna o JSON formatado
        return json_encode($result, JSON_UNESCAPED_UNICODE | JSON_PRETTY_PRINT);
    }

    private function arrayvalues(array &$result, $grupos): void
    {
        $ref = &$result;
        // remove o primero grupo
        $sGrupo = array_shift($grupos);
        foreach ($ref as $key => &$assoc) {
            if ($sGrupo == $key) {
                $ref[$key] = array_values($assoc);
                if (!$grupos) continue;
                foreach ($ref[$key] as &$reg) {
                    $this->arrayvalues($reg, $grupos);
                }
            }
        }
    }
    
    /**
     * Agrupa um registro em $result conforme os grupos definidos
     * @param array &$result
     * @param array $assoc
     * @param array $grupos
     */
    private function addToGroup(array &$result, array $assoc, array $grupos): void
    {
        $ref = &$result;
        foreach ($grupos as $grupo => $aCamposGrupo) {
            $key = $assoc[$aCamposGrupo[0]] ?? '';
            if (!isset($ref[$grupo])) $ref[$grupo] = [];
            $ref = &$ref[$grupo];
            if (!isset($ref[$key])) {
                // Só os campos do grupo
                $ref[$key] = [];
                foreach ($aCamposGrupo as $campo) {
                    $ref[$key][$campo] = $assoc[$campo] ?? null;
                    unset($assoc[$campo]);
                }
            } else {
                // Mesmo assim, remove os campos do grupo do assoc
                foreach ($aCamposGrupo as $campo) {
                    unset($assoc[$campo]);
                }
            }
            $ref = &$ref[$key];
        }
        $ref['item'] = $assoc;
    }
}
