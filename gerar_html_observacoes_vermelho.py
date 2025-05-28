
import pandas as pd

def gerar_html(excel_path, output_html):
    df = pd.read_excel(excel_path)
    df.columns = [col.strip() for col in df.columns]

    df = df[
        ["Prefixo e Sufixo do Imóvel?",
         "Gestão e Loja/Equipa do Ref Imóvel EG***RF?",
         "Responsável (atribuir)",
         "Angariador Inactivo (atribuir)",
         "Imóvel Exclusivo",
         "Imóvel Aberto",
         "Angariador Não Atende",
         "Equipa (atribuir",
         "quem pode ver/editar",
         "Observações",
         "Contacto e Horário da Loja"]
    ]
    df.columns = ["prefix", "loja", "responsavel", "inactivo", "exclusivo", "aberto",
                  "nao_atende", "equipa", "editar", "observacoes", "contacto"]
    df = df.fillna("")
    df = df.drop_duplicates()

    items = []
    for _, row in df.iterrows():
        item = (
            f'{{ prefix: "{row.prefix}", loja: "{row.loja}", responsavel: "{row.responsavel}", '
            f'inactivo: "{row.inactivo}", exclusivo: "{row.exclusivo}", aberto: "{row.aberto}", '
            f'nao_atende: "{row.nao_atende}", equipa: "{row.equipa}", editar: "{row.editar}", '
            f'observacoes: "{row.observacoes}", contacto: "{row.contacto}" }}'
        )
        items.append(item)
    js_array = ",\n      ".join(items)

    html = f"""<!DOCTYPE html>
<html lang="pt">
<head>
  <meta charset="UTF-8" />
  <meta name="viewport" content="width=device-width, initial-scale=1.0"/>
  <title>Mapeamento de Referências</title>
  <style>
    body {{ font-family: Arial, sans-serif; max-width: 900px; margin: 40px auto; padding: 20px; }}
    img.logo {{ max-width: 180px; display: block; margin: 0 auto 20px; }}
    input {{ width: 100%; padding: 10px; font-size: 16px; margin-top: 10px; margin-bottom: 20px; }}
    .card {{ border: 1px solid #ccc; border-radius: 8px; padding: 20px; background-color: #f9f9f9; }}
    .label {{ font-weight: bold; margin-top: 10px; }}
    .error {{ color: red; font-size: 14px; }}
  </style>
</head>
<body>
  <img src="logo_easygest.jpg" alt="EasyGest" class="logo" />
  <h1>Mapeamento de Referências</h1>
  <label for="ref">Referência do Imóvel:</label>
  <input type="text" id="ref" placeholder="Ex: EG123RF" oninput="checkReference()" />

  <div id="result"></div>

  <script>
    const data = [
      {js_array}
    ];

    function normalizar(ref) {{
      return ref.toUpperCase().replace(/\s+/g, '');
    }}

    function checkReference() {{
      const input = normalizar(document.getElementById('ref').value);
      const resultDiv = document.getElementById('result');
      resultDiv.innerHTML = "";

      if (input.length < 3) return;

      const match = data.find(entry => {{
        const prefixParts = entry.prefix.split('***');
        const startsOk = input.startsWith(prefixParts[0]);
        const endsOk = prefixParts[1] ? input.endsWith(prefixParts[1]) : true;
        return startsOk && endsOk;
      }});

      if (match) {{
        resultDiv.innerHTML = `
          <div class="card">
            <p><span class="label">Loja/Equipa:</span> ${{match.loja}}</p>
            <p><span class="label">Responsável (atribuir):</span> ${{match.responsavel}}</p>
            <p><span class="label">Angariador Inactivo:</span> ${{match.inactivo}}</p>
            <p><span class="label">Imóvel Exclusivo:</span> ${{match.exclusivo}}</p>
            <p><span class="label">Imóvel Aberto:</span> ${{match.aberto}}</p>
            <p><span class="label">Angariador Não Atende:</span> ${{match.nao_atende}}</p>
            <p><span class="label">Equipa (atribuir):</span> ${{match.equipa}}</p>
            <p><span class="label">Quem pode ver/editar:</span> ${{match.editar}}</p>
            <p><span class="label">Observações:</span> <span style="color:#b30000; font-weight:bold;">${{match.observacoes}}</span></p>
            <p><span class="label">Contacto e Horário da Loja:</span> ${{match.contacto}}</p>
          </div>
        `;
      }} else {{
        resultDiv.innerHTML = '<p class="error">Referência não encontrada.</p>';
      }}
    }}
  </script>
</body>
</html>
"""

    with open(output_html, "w", encoding="utf-8") as f:
        f.write(html)

if __name__ == "__main__":
    gerar_html("Mapeamento ImoEasygest.xlsx", "ref_mapping_com_logo.html")
    print("✔ HTML gerado com 11 colunas.")
