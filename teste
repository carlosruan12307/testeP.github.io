

  (function() {
    const dominioPermitido = "seu-moodle.com.br";
    const referrer = document.referrer;

    // Se não houver referrer ou se o domínio não for o do Moodle
    if (!referrer || !referrer.includes(dominioPermitido)) {
        alert("Acesso negado: Este recurso só pode ser visualizado através do Moodle.");
        window.location.href = "https://google.com"; // Redireciona o "curioso"
    }
})();
  function sanitizeSheetName(name) {
    return name.replace(/[:\\\/\?\*\[\]]/g, '-').substring(0, 31);
  }
  
  async function executarTodos() {
    const response = await fetch("https://ead.unifor.br/ava/mod/resource/view.php?id=4578168");
    const ResponseConfig = await response.json()
    const linksAgendamento = ResponseConfig["100%"]["1semestre_A"]["relatorioAgendamento_1chamada"].urls

    const workbookFinal = XLSX.utils.book_new();
  
    const progressContainer = document.getElementById('progressContainer');
    const progressFill = document.getElementById('progressFill');
    const progressPercent = document.getElementById('progressPercent');
    const progressText = document.getElementById('progressText');
  
    if (progressContainer) progressContainer.style.display = 'flex';
  
    for (let i = 0; i < linksAgendamento.length; i++) {
      const url = `${linksAgendamento[i]}&download=xls&sesskey=${getSesskey()}`;
      const responsePagina = await fetch(linksAgendamento[i]);
      const html = await responsePagina.text();
      
      const parser = new DOMParser();
      const doc = parser.parseFromString(html, 'text/html');
      
      // Tenta capturar o título da página
      let nomeRecurso = doc.querySelector(".page-header-headings h1").textContent
      
      // Sanitiza o nome para aba do Excel
      nomeRecurso = sanitizeSheetName(nomeRecurso);
      if (progressText) progressText.textContent = `Baixando arquivo ${i + 1} de ${linksAgendamento.length}...`;
      if (progressPercent) progressPercent.textContent = `${Math.round((i / linksAgendamento.length) * 100)}%`;
      if (progressFill) progressFill.style.width = `${(i / linksAgendamento.length) * 100}%`;
  
      try {
        const response = await fetch(url);
  
        if (!response.ok) {
          console.error(`Erro ao baixar o arquivo ${url}: status ${response.status}`);
          continue;
        }
  
        const contentType = response.headers.get("content-type") || "";
  
        if (!contentType.includes("application/vnd.ms-excel") &&
            !contentType.includes("application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")) {
          console.error(`Arquivo não é Excel. Content-Type: ${contentType}`);
          const text = await response.text();
          console.log("Conteúdo recebido (início):", text.substring(0, 500));
          continue;
        }
  
        const arrayBuffer = await response.arrayBuffer();
        const workbook = XLSX.read(arrayBuffer, { type: "array" });
  
        workbook.SheetNames.forEach(sheetName => {
          const sheet = workbook.Sheets[sheetName];
          const data = XLSX.utils.sheet_to_json(sheet, { defval: "" });
  
          let abaNome = sheetName;
  
          if (data.length > 0) {
            const primeiraLinha = data[0];
            const escolha = primeiraLinha["Escolha {$a}"];
            if (escolha) {
              const match = escolha.match(/^(\d{2}\/\d{2})/);
              if (match) {
                abaNome = sanitizeSheetName(match[1].replace('/', '-')); // substitui / por -
              }
            }
          }
  
          // Sanitize e evitar duplicata
        
          let newSheetName = nomeRecurso;
          
  
          XLSX.utils.book_append_sheet(workbookFinal, sheet, newSheetName);
        });
  
      } catch (e) {
        console.error("Erro no fetch ou processamento:", e);
      }
    }
  
    if (workbookFinal.SheetNames.length === 0) {
      console.error("Nenhuma planilha válida para gerar arquivo final.");
      if (progressText) progressText.textContent = "Nenhum arquivo válido baixado.";
      return;
    }
  
    if (progressText) progressText.textContent = `Finalizando...`;
    if (progressPercent) progressPercent.textContent = `100%`;
    if (progressFill) progressFill.style.width = `100%`;
  
    const wbout = XLSX.write(workbookFinal, { bookType: "xlsx", type: "array" });
    const blob = new Blob([wbout], { type: "application/octet-stream" });
    const urlBlob = window.URL.createObjectURL(blob);
    const a = document.createElement("a");
    a.href = urlBlob;
    a.download = "relatorio_unificado.xlsx";
    document.body.appendChild(a);
    a.click();
    a.remove();
    window.URL.revokeObjectURL(urlBlob);
  
    if (progressText) progressText.textContent = "Download concluído!";
    if (progressContainer) progressContainer.style.display = 'none';
}
  
  function getSesskey() {
    if (window.M && M.cfg && M.cfg.sesskey) {
      return M.cfg.sesskey;
    }
    const inputSesskey = document.querySelector('input[name="sesskey"]');
    if (inputSesskey) {
      return inputSesskey.value;
    }
    return null;
  }
  
  executarTodos();
