"""
gerar_relacoes.py
=================
Extrai combinações únicas de dimensões (CANAL × SEGMENTO × EMPRESA × VENDEDOR × GRUPO)
do fFaturamento via Analysis Services local do Power BI Desktop.
Gera dim_relacoes.json localmente e publica no OneDrive via N8N.

Pré-requisito
    Power BI Desktop aberto com Comercial.pbip (modelo carregado e atualizado)

Uso
    python gerar_relacoes.py
"""

import os, json, sys, subprocess, tempfile, urllib.request, urllib.parse
from datetime import datetime

SCRIPT_DIR    = os.path.dirname(os.path.abspath(__file__))
OUTPUT_JSON   = os.path.join(SCRIPT_DIR, "dim_relacoes.json")
N8N_WRITE_URL = "https://n8n.xperiun.com/webhook/ibratin-campanhas"

# Importa utilitários do script de dimensões existente (porta AS + caminho ADOMD)
sys.path.insert(0, SCRIPT_DIR)
from exportar_dims_pbi import encontrar_porta, encontrar_adomd

# ---------------------------------------------------------------------------
# Template PowerShell – retorna JSON das combinações únicas via DAX
# ---------------------------------------------------------------------------
_PS = r"""
param([string]$Port, [string]$AdomdDll, [string]$OutFile)
$ErrorActionPreference = 'Stop'

Add-Type -Path $AdomdDll
$conn = New-Object Microsoft.AnalysisServices.AdomdClient.AdomdConnection(
    "Data Source=localhost:$Port;Timeout=600")
$conn.Open()

function RunDAX([string]$Dax) {
    $cmd = $conn.CreateCommand()
    $cmd.CommandText = $Dax
    $reader = $cmd.ExecuteReader()
    $n      = $reader.FieldCount
    $cols   = @(); for ($i = 0; $i -lt $n; $i++) { $cols += $reader.GetName($i) }
    $rows   = [System.Collections.Generic.List[object]]::new()
    while ($reader.Read()) {
        $row = [ordered]@{}
        for ($i = 0; $i -lt $n; $i++) {
            $v         = $reader.GetValue($i)
            $row[$cols[$i]] = if ($null -eq $v) { '' } else { "$v".Trim() }
        }
        $rows.Add([pscustomobject]$row)
    }
    $reader.Close()
    return ,$rows
}
function TentarDAX([string]$Dax) { try { return RunDAX($Dax) } catch { return @() } }

# Tentativa 1: SUMMARIZECOLUMNS com dProdutos (via relacionamento do modelo)
$rows = TentarDAX("
EVALUATE
SUMMARIZECOLUMNS(
    fFaturamento[CANAL],
    fFaturamento[SEGMENTO],
    fFaturamento[EMPRESA],
    fFaturamento[VENDEDOR],
    dProdutos[Grupo Produto]
)
")

# Tentativa 2: apenas colunas do fFaturamento (sem dProdutos)
if ($rows.Count -eq 0) {
    $rows = TentarDAX("
EVALUATE
SUMMARIZECOLUMNS(
    fFaturamento[CANAL],
    fFaturamento[SEGMENTO],
    fFaturamento[EMPRESA],
    fFaturamento[VENDEDOR]
)
")
}

$conn.Close()

# Serializa para JSON compacto
@{ total = $rows.Count; rows = $rows } |
    ConvertTo-Json -Depth 3 -Compress |
    Set-Content -Path $OutFile -Encoding UTF8

Write-Host "OK:$($rows.Count)"
"""


def _normalizar_col(nome: str) -> str:
    """'[CANAL]' → 'CANAL', 'dProdutos[Grupo Produto]' → 'Grupo Produto'"""
    if "[" in nome:
        return nome.split("[")[-1].rstrip("]").strip()
    return nome.strip()


# Matchers: nome normalizado → chave canônica
_MATCHERS = {
    "canal":    lambda n: n.upper() == "CANAL",
    "segmento": lambda n: n.upper() == "SEGMENTO",
    "empresa":  lambda n: n.upper() in ("EMPRESA", "NOME_EMPRESA", "NM_EMPRESA"),
    "vendedor": lambda n: n.upper() in ("VENDEDOR", "NOME_VENDEDOR", "NOMEVEND", "NM_VENDEDOR"),
    "grupo":    lambda n: "grupo" in n.lower(),
}


def _mapear_colunas(primeiro_row: dict) -> dict:
    """Retorna {col_original: chave_canônica} para as colunas reconhecidas."""
    mapa = {}
    for col_raw in primeiro_row.keys():
        norm = _normalizar_col(col_raw)
        for chave, matcher in _MATCHERS.items():
            if chave not in mapa.values() and matcher(norm):
                mapa[col_raw] = chave
                break
    return mapa


def executar_relacoes(porta: str, adomd: str) -> list:
    ps_file  = os.path.join(tempfile.gettempdir(), "_pbi_relacoes.ps1")
    raw_json = os.path.join(tempfile.gettempdir(), "_pbi_relacoes.json")

    with open(ps_file, "w", encoding="utf-8") as f:
        f.write(_PS)

    result = subprocess.run(
        ["powershell", "-NonInteractive", "-ExecutionPolicy", "Bypass",
         "-File", ps_file, "-Port", porta, "-AdomdDll", adomd, "-OutFile", raw_json],
        capture_output=True, text=True, encoding="utf-8", timeout=600,
    )

    stderr = result.stderr.strip()
    stdout = result.stdout.strip()

    if result.returncode != 0 and "OK:" not in stdout:
        raise RuntimeError(stderr or f"PowerShell saiu com código {result.returncode}")

    if not os.path.isfile(raw_json):
        raise RuntimeError("Arquivo JSON não foi gerado pelo PowerShell")

    with open(raw_json, encoding="utf-8-sig") as f:
        payload = json.load(f)

    return payload.get("rows") or []


def processar_tuplas(raw_rows: list) -> list:
    if not raw_rows:
        return []

    mapa = _mapear_colunas(raw_rows[0])
    if not mapa:
        raise RuntimeError(
            "Colunas não reconhecidas: " + str(list(raw_rows[0].keys())) +
            "\nEspera CANAL, SEGMENTO, EMPRESA, VENDEDOR e/ou Grupo Produto"
        )

    tuplas = []
    for row in raw_rows:
        t = {chave: str(row.get(col_raw, "")).strip()
             for col_raw, chave in mapa.items()
             if str(row.get(col_raw, "")).strip()}
        if t:
            tuplas.append(t)

    return tuplas


def upload_onedrive(content: str, arquivo: str) -> int:
    data = json.dumps({"arquivo": arquivo, "conteudo": content, "formato": "plain"}).encode("utf-8")
    req  = urllib.request.Request(
        N8N_WRITE_URL, data=data,
        headers={"Content-Type": "application/json"}, method="POST"
    )
    with urllib.request.urlopen(req, timeout=60) as resp:
        return resp.status


def main():
    print("=" * 62)
    print("  Gerar Relações de Dimensões — Power BI Desktop (AS local)")
    print("  Ibratin Tintas e Texturas")
    print("=" * 62)

    # 1. ADOMD.NET
    print("\n[1/5] Localizando ADOMD.NET …")
    adomd = encontrar_adomd()
    if not adomd:
        print("  [ERRO] ADOMD.NET não encontrado.")
        print("  Instale as Client Libraries em:")
        print("  https://learn.microsoft.com/pt-br/analysis-services/client-libraries")
        sys.exit(1)
    print(f"  OK: {adomd}")

    # 2. Porta AS local
    print("\n[2/5] Localizando Power BI Desktop (porta AS local) …")
    porta = encontrar_porta()
    if not porta:
        print("  [ERRO] Power BI Desktop não encontrado ou não está com dados carregados.")
        print("  Abra Comercial.pbip e aguarde o carregamento completo.")
        sys.exit(1)
    print(f"  Porta AS: {porta}")

    # 3. Executar DAX
    print("\n[3/5] Extraindo combinações de dimensões via SUMMARIZECOLUMNS …")
    try:
        raw_rows = executar_relacoes(porta, adomd)
    except (RuntimeError, subprocess.TimeoutExpired, OSError) as e:
        print(f"  [ERRO] {e}")
        sys.exit(1)

    if not raw_rows:
        print("  [AVISO] Nenhuma linha retornada.")
        print("  Verifique se Comercial.pbip está aberto e com dados atualizados.")
        sys.exit(1)

    print(f"  {len(raw_rows):,} combinações únicas encontradas")

    # 4. Processar e salvar
    print("\n[4/5] Processando e salvando dim_relacoes.json …")
    try:
        tuplas = processar_tuplas(raw_rows)
    except RuntimeError as e:
        print(f"  [ERRO] {e}")
        sys.exit(1)

    relacoes = {
        "_meta": {
            "gerado": datetime.now().strftime("%Y-%m-%dT%H:%M:%S"),
            "total": len(tuplas),
        },
        "tuplas": tuplas,
    }
    content_str = json.dumps(relacoes, ensure_ascii=False)

    with open(OUTPUT_JSON, "w", encoding="utf-8") as f:
        f.write(content_str)

    size_kb = len(content_str.encode("utf-8")) / 1024
    print(f"  Salvo: {OUTPUT_JSON}")
    print(f"  Tamanho: {size_kb:.0f} KB | {len(tuplas):,} tuplas")

    # 5. Upload OneDrive
    print("\n[5/5] Publicando no OneDrive via N8N …")
    try:
        status = upload_onedrive(content_str, "dim_relacoes.json")
        print(f"  OK — HTTP {status}")
    except Exception as e:
        print(f"  [AVISO] Upload falhou: {e}")
        print("  O arquivo foi salvo localmente em dim_relacoes.json.")
        print("  Publique manualmente se necessário.")

    # Resumo
    print("\n" + "=" * 62)
    print("  CONCLUÍDO")
    print("=" * 62)

    canais    = sorted({t["canal"]    for t in tuplas if "canal"    in t})
    segmentos = sorted({t["segmento"] for t in tuplas if "segmento" in t})
    empresas  = sorted({t["empresa"]  for t in tuplas if "empresa"  in t})
    vendedores = sorted({t["vendedor"] for t in tuplas if "vendedor" in t})
    grupos    = sorted({t["grupo"]    for t in tuplas if "grupo"    in t})

    def resumir(lst, n=5):
        s = ", ".join(lst[:n])
        return s + (" …" if len(lst) > n else "")

    print(f"\n  Canais    ({len(canais):>3}): {resumir(canais)}")
    print(f"  Segmentos ({len(segmentos):>3}): {resumir(segmentos)}")
    print(f"  Empresas  ({len(empresas):>3}): {resumir(empresas)}")
    print(f"  Vendedores({len(vendedores):>3}): {resumir(vendedores)}")
    print(f"  Grupos    ({len(grupos):>3}): {resumir(grupos)}")
    print()
    print("  Próximo passo:")
    print("    Recarregue o gestor-campanhas.html (F5) ou clique em")
    print("    'Relações PBI' → OK no gestor para aplicar os filtros.")


if __name__ == "__main__":
    main()
