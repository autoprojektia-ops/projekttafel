# Gera o estudo de caso em Word (tabelas limpas + negritos do PDF)
$ErrorActionPreference = 'Stop'
$out = Join-Path $PSScriptRoot 'Estudo-de-caso-Guaras-COM-SUMARIO.docx'
if (Test-Path -LiteralPath $out) { Remove-Item -LiteralPath $out -Force }

Get-Process WINWORD -ErrorAction SilentlyContinue | Stop-Process -Force -ErrorAction SilentlyContinue
Start-Sleep -Seconds 1

$word = New-Object -ComObject Word.Application
$word.Visible = $false
$word.DisplayAlerts = 0
$doc = $word.Documents.Add()
$sel = $word.Selection

# Pagina A4, margens confortaveis
$doc.PageSetup.TopMargin = $word.CentimetersToPoints(2)
$doc.PageSetup.BottomMargin = $word.CentimetersToPoints(2)
$doc.PageSetup.LeftMargin = $word.CentimetersToPoints(2)
$doc.PageSetup.RightMargin = $word.CentimetersToPoints(2)

$script:White = 16777215
$script:Black = 0
$script:Ink = (36 + (48 * 256) + (68 * 65536))      # #243044
$script:Mute = (85 + (96 * 256) + (112 * 65536))     # #556070
$script:Signal = (196 + (92 * 256) + (42 * 65536))   # #c45c2a
$script:HeaderBg = (61 + (90 * 256) + (128 * 65536)) # #3d5a80
$script:Zebra = (247 + (249 * 256) + (251 * 65536))  # #f7f9fb
$script:Line = (201 + (210 * 256) + (220 * 65536))   # #c9d2dc

function Set-Run([float]$size = 11, [bool]$bold = $false, $color = $null) {
  if ($null -eq $color) { $color = $script:Ink }
  $sel.Font.Name = 'Calibri'
  $sel.Font.Size = $size
  $sel.Font.Bold = $bold
  $sel.Font.Color = $color
}

function P-Reset([int]$after = 8, [int]$before = 0) {
  $sel.ParagraphFormat.SpaceAfter = $after
  $sel.ParagraphFormat.SpaceBefore = $before
  $sel.ParagraphFormat.Alignment = 0
  $sel.ParagraphFormat.LineSpacingRule = 0
}

function Write-Plain([string]$text, [float]$size = 11, [bool]$bold = $false, $color = $null, [int]$after = 8) {
  P-Reset $after 0
  Set-Run $size $bold $color
  $sel.TypeText($text)
  $sel.TypeParagraph()
}

function Write-Mixed([object[]]$parts, [float]$size = 11, [int]$after = 8) {
  P-Reset $after 0
  foreach ($p in $parts) {
    $b = $false
    if ($p.ContainsKey('b')) { $b = [bool]$p.b }
    $c = $script:Ink
    if ($p.ContainsKey('c')) { $c = $p.c }
    Set-Run $size $b $c
    $sel.TypeText([string]$p.t)
  }
  $sel.TypeParagraph()
}

function Write-H2([string]$text) {
  # Estilo Titulo 2 = aparece no sumario do Google Docs
  $sel.Style = $doc.Styles.Item(-3)
  $sel.Font.Name = "Calibri"
  $sel.Font.Size = 13
  $sel.Font.Bold = $true
  $sel.Font.Color = $script:Ink
  $sel.ParagraphFormat.SpaceBefore = 16
  $sel.ParagraphFormat.SpaceAfter = 10
  $sel.TypeText($text)
  $sel.TypeParagraph()
  $sel.Style = $doc.Styles.Item(-1)
}

function Set-CellText($cell, [string]$text, [bool]$bold, $fg, $bg, [float]$size = 10, [bool]$center = $false) {
  # Limpa e escreve sem o marcador da celula estragar o estilo
  $rng = $cell.Range
  $rng.Text = $text
  # Exclui a marca final da celula
  $rng.MoveEnd(1, -1) | Out-Null
  $rng.Font.Name = 'Calibri'
  $rng.Font.Size = $size
  $rng.Font.Bold = $bold
  $rng.Font.Color = $fg
  $rng.ParagraphFormat.SpaceAfter = 3
  $rng.ParagraphFormat.SpaceBefore = 3
  if ($center) { $rng.ParagraphFormat.Alignment = 1 } else { $rng.ParagraphFormat.Alignment = 0 }
  $cell.Shading.BackgroundPatternColor = $bg
  $cell.VerticalAlignment = 1
  $cell.TopPadding = 4
  $cell.BottomPadding = 4
  $cell.LeftPadding = 5
  $cell.RightPadding = 5
}

function Style-Borders($table) {
  $borders = @(1,2,3,4,5,6) # Outside + Inside H/V
  foreach ($b in $borders) {
    try {
      $border = $table.Borders.Item($b)
      $border.LineStyle = 1
      $border.LineWidth = 4 # 0.5 pt
      $border.Color = $script:Line
    } catch {}
  }
}

function Add-NiceTable([string[][]]$rows, [double[]]$widthsCm, [int[]]$centerCols = @()) {
  $nRows = $rows.Count
  $nCols = $rows[0].Count
  $table = $doc.Tables.Add($sel.Range, $nRows, $nCols)
  $table.AllowAutoFit = $false
  Style-Borders $table

  # Larguras fixas em cm
  if ($widthsCm -and $widthsCm.Count -eq $nCols) {
    for ($c = 0; $c -lt $nCols; $c++) {
      $table.Columns.Item($c + 1).PreferredWidthType = 3 # wdPreferredWidthPoints? 3=cm in Word
      $table.Columns.Item($c + 1).PreferredWidth = $widthsCm[$c]
      $table.Columns.Item($c + 1).Width = $word.CentimetersToPoints($widthsCm[$c])
    }
  } else {
    $table.AutoFitBehavior(2)
  }

  for ($r = 0; $r -lt $nRows; $r++) {
    for ($c = 0; $c -lt $nCols; $c++) {
      $isHeader = ($r -eq 0)
      $bg = if ($isHeader) { $script:HeaderBg } elseif (($r % 2) -eq 0) { $script:Zebra } else { $script:White }
      $fg = if ($isHeader) { $script:White } else { $script:Black }
      $center = ($centerCols -contains ($c + 1))
      Set-CellText $table.Cell($r + 1, $c + 1) ([string]$rows[$r][$c]) $isHeader $fg $bg 10 $center
    }
  }

  $sel.EndKey(6) | Out-Null
  $sel.TypeParagraph()
}

# ========== CONTEUDO ==========

P-Reset 6 0
Set-Run 10 $true $script:Signal
$sel.TypeText('Trilhas 2026 · Desafio 01 · estudo de caso')
$sel.TypeParagraph()

$sel.Style = $doc.Styles.Item(-2)
$sel.Font.Name = "Georgia"
$sel.Font.Size = 18
$sel.Font.Bold = $true
$sel.Font.Color = $script:Ink
$sel.ParagraphFormat.SpaceAfter = 8
$sel.TypeText("Liberacao de colaborador para nova obra apos conferencia dos treinamentos obrigatorios")
$sel.TypeParagraph()
$sel.Style = $doc.Styles.Item(-1)

Write-Plain 'Manutencao Guaras · caso Rafael Mendes, matricula 4587 · 5 e 6 de outubro' 11 $false $script:Mute 6
Write-Plain 'Documento unico da entrega. A analise parte so do briefing oficial da empresa. Contratacao, aplicacao de treinamento e gestao de certificados ficam fora deste processo.' 10 $false $script:Mute 10

Write-Plain 'O que este documento responde' 11 $true $script:Ink 4
Write-Plain 'Mapa AS-IS, tempos, gargalo, 5 Porques, direcao sem ferramenta, 3 partes, algoritmo com decisao, gatilho / condicao / acao, e a automacao Make ligada ao gargalo.' 11 $false $script:Ink 12

Write-H2 'Empresa, processo e ferramentas atuais'
Write-Plain 'A Manutencao Guaras presta manutencao industrial em varias obras ao mesmo tempo. Antes de um colaborador ir para uma obra nova, a empresa confirma se os treinamentos obrigatorios daquela obra estao validos. So entao a mobilizacao e liberada.'
Write-Mixed @(
  @{ t = 'Entra:'; b = $true },
  @{ t = ' pedido do supervisor (nome, matricula, funcao, obra de destino). ' },
  @{ t = 'Sai:'; b = $true },
  @{ t = ' confirmacao de que a pessoa esta liberada para a obra.' }
)
Write-Plain 'Ferramentas usadas hoje: e-mail, planilha de controle de mobilizacoes e planilha com os treinamentos dos colaboradores.'

Add-NiceTable @(
  @('Papel', 'O que faz neste processo'),
  @('Supervisor da obra', 'Pede a liberacao e espera a confirmacao.'),
  @('Assistente de mobilizacao', 'Registra o pedido na planilha e, no fim, marca "Liberado para a obra" e avisa o supervisor.'),
  @('Responsavel pelos treinamentos', 'Confere na planilha de treinamentos se os 3 cursos daquela obra estao validos e atualiza o estado.'),
  @('Coordenador de Seguranca', 'Valida quem esta "Apto para validacao" e avisa a equipe de mobilizacao.')
) @(5.5, 11.5)

Write-Plain 'Regra da empresa: se os 3 treinamentos estao validos, a pessoa pode ir para validacao. Se algum esta pendente ou vencido, a regularizacao fica fora deste processo — aqui so segue quem ja pode ser validado.'

Write-H2 'Caso observado — Rafael Mendes'
Write-Plain 'No dia 5 de outubro o supervisor pediu o mecanico Rafael Mendes, matricula 4587, para a Unidade Industrial Norte. Os 3 treinamentos obrigatorios dessa obra estavam validos: Integracao de Seguranca da Obra, Trabalho em Altura e Seguranca em Ferramentas Eletricas. O pedido foi encerrado no dia seguinte, as 10:34, com o estado "Liberado para a obra".'

Add-NiceTable @(
  @('Quando', 'O que aconteceu', 'Trabalho'),
  @('05/10 08:00-08:06', 'Supervisor prepara e envia o e-mail (nome, matricula, funcao, obra).', '6 min'),
  @('05/10 08:06-09:00', 'E-mail ja enviado. Assistente ainda nao registrou.', 'espera'),
  @('05/10 09:00-09:08', 'Assistente registra na planilha e marca "aguardando conferencia dos treinamentos".', '8 min'),
  @('05/10 09:08-15:00', 'Pedido na planilha. Equipe de treinamentos so confere duas vezes ao dia.', 'espera'),
  @('05/10 15:00-15:14', 'Conferencia dos 3 cursos. Todos validos. Estado "Apto para validacao". Nenhum recado e enviado.', '14 min'),
  @('05/10 15:14 -> 06/10 09:00', 'Coordenador nao sabe que o Rafael esta esperando. So abre a planilha de manha.', 'espera'),
  @('06/10 09:00-09:05', 'Coordenador ve o Rafael, valida e envia e-mail a equipe de mobilizacao.', '5 min'),
  @('06/10 09:05-10:30', 'Aprovacao ja enviada. Assistente ainda nao retomou o pedido.', 'espera'),
  @('06/10 10:30-10:34', 'Assistente marca "Liberado para a obra" e confirma ao supervisor por e-mail.', '4 min')
) @(4.2, 10.3, 2.5) @(3)

Write-Plain 'Entrevistas do briefing: nao existe aviso definido quando a solicitacao muda de etapa. O controle oficial e a planilha. O coordenador disse que o problema e saber quando apareceu alguem novo — em geral ele verifica de manha.'

Write-H2 '1. Mapa AS-IS'
Write-Plain 'Cinco etapas. Em cada linha: quem faz, tempo de trabalho e espera ate a etapa seguinte comecar. Um dia = 24 horas de calendario.'

Add-NiceTable @(
  @('Etapa', 'O que acontece', 'Trabalho', 'Espera'),
  @('1', 'Supervisor solicita por e-mail', '6 min', '54 min'),
  @('2', 'Assistente registra na planilha', '8 min', '5 h 52 min'),
  @('3 (gargalo)', 'Treinamentos conferem os 3 cursos (sem aviso)', '14 min', '17 h 46 min'),
  @('4', 'Coordenador valida', '5 min', '1 h 25 min'),
  @('5', 'Assistente libera e confirma', '4 min', 'Fim 10:34')
) @(2.8, 8.5, 2.5, 3.2) @(1,3,4)

Write-Mixed @(
  @{ t = 'Lead time: '; b = $true },
  @{ t = '26 h 34 min  ·  ' },
  @{ t = 'Trabalho: '; b = $true },
  @{ t = '37 min  ·  ' },
  @{ t = 'Espera: '; b = $true },
  @{ t = '25 h 57 min  ·  ' },
  @{ t = 'Gargalo: '; b = $true },
  @{ t = '67%' }
) 11 6
Write-Plain 'Inicio: 05/10 08:00. Fim: 06/10 10:34. Trabalho 6+8+14+5+4 = 37 min. Esperas 54 min + 5 h 52 min + 17 h 46 min + 1 h 25 min = 25 h 57 min.' 10 $false $script:Mute 10

Write-H2 '2. Tempos de trabalho, espera e lead time'
Add-NiceTable @(
  @('Metrica', 'No caso Rafael', 'Em horas (1 dia = 24 h)'),
  @('Trabalho (valor)', '37 min', '0,62 h'),
  @('Espera (desperdicio)', '25 h 57 min', '25,95 h'),
  @('Lead time (trabalho + espera)', '26 h 34 min', '26,57 h')
) @(7.0, 5.0, 5.0) @(2,3)
Write-Plain 'Quase todo o lead time e espera. O trabalho de cada pessoa e curto. Ninguem "demora a fazer". O pedido para entre uma etapa e a seguinte.'

Write-H2 '3. Gargalo'
Write-Plain 'Gargalo = a maior espera que se repete no modo de trabalhar, nao um acaso daquele dia.'
Add-NiceTable @(
  @('Espera depois da etapa', 'Tempo', '%', 'Por que se repete'),
  @('1. E-mail do supervisor ainda nao visto', '54 min', '3%', 'Assistente ve o e-mail quando puder.'),
  @('2. Pedido esperando a conferencia', '5 h 52 min', '22%', 'Treinamentos so conferem duas vezes ao dia.'),
  @('3. "Apto para validacao" sem recado', '17 h 46 min', '67%', 'Ninguem avisa o coordenador. Ele abre a planilha de manha.'),
  @('4. E-mail de aprovacao ainda nao visto', '1 h 25 min', '5%', 'Assistente retoma quando ve o e-mail.')
) @(6.5, 2.8, 1.5, 6.2) @(2,3)

Write-Plain 'Gargalo' 11 $true $script:Signal 4
Write-Mixed @(
  @{ t = 'Depois que os treinamentos marcam "Apto para validacao", o pedido espera 17 h 46 min (17,77 h). Percentual: (17,77 / 26,57) x 100 = ' },
  @{ t = '67%'; b = $true },
  @{ t = '. A analise do coordenador durou 5 minutos. A espera nao e o trabalho dele — e ele nao saber que o Rafael ja estava na fila.' }
) 11 8
Write-Plain 'A segunda espera (5 h 52 min, 22%) tambem e padrao (conferencia so duas vezes ao dia). Fica atras do gargalo.'

Write-H2 '4. Cinco Porques e tipo de espera'
Write-Mixed @(
  @{ t = '1. ' },
  @{ t = 'Por que'; b = $true },
  @{ t = ' o pedido ficou 17 h 46 min depois de "Apto para validacao"? Porque o coordenador de seguranca nao soube que o Rafael estava esperando. Ninguem avisou.' }
)
Write-Mixed @(
  @{ t = '2. ' },
  @{ t = 'Por que'; b = $true },
  @{ t = ' ele nao soube? Porque, quando os treinamentos atualizam o estado na planilha, nao parte nenhum recado. Ele so abre a planilha de manha.' }
)
Write-Mixed @(
  @{ t = '3. ' },
  @{ t = 'Por que'; b = $true },
  @{ t = ' nao parte recado? Porque o controle oficial e a planilha: quem termina a etapa atualiza o estado e considera o trabalho feito.' }
)
Write-Mixed @(
  @{ t = '4. ' },
  @{ t = 'Por que'; b = $true },
  @{ t = ' a etapa seguinte nao comeca na hora? Porque olhar a planilha nao e disparado pela mudanca de estado — depende do horario de costume de cada um.' }
)
Write-Mixed @(
  @{ t = '5. ' },
  @{ t = 'Por que'; b = $true },
  @{ t = ' nunca houve outro modo? Porque a empresa nunca definiu um aviso quando a solicitacao muda de etapa. O gestor disse isso na entrevista.' }
)
Write-Mixed @(
  @{ t = 'Causa raiz:'; b = $true },
  @{ t = ' falta um aviso quando a solicitacao muda de etapa.' }
)
Write-Mixed @(
  @{ t = 'Tipo de espera:'; b = $true },
  @{ t = ' informacao. Quem precisa agir nao recebe, no momento certo, o dado de que ja pode agir. Nao e falta de gente na etapa 4 (5 minutos de trabalho) e nao e material.' }
)

Write-H2 '5. Direcao da solucao (ainda sem ferramenta)'
Write-Plain 'Quando uma solicitacao mudar de etapa, quem precisa agir na seguinte recebe automaticamente um aviso com uma versao visivel do pedido — sem depender de abrir a planilha no horario de costume.' 11 $true $script:Ink 8
Write-Plain 'Isso ataca o gargalo (coordenador nao sabe que ficou Apto) e tambem a espera menor depois da validacao (assistente nao sabe que ja pode liberar). Nao se escolhe ainda e-mail automatico, planilha nova ou outra ferramenta: so a regra do fluxo.'

Write-H2 '6. Decomposicao (pensamento computacional)'
Write-Plain 'A direcao parte em 3 pedacos. Cada um da para testar sozinho. Juntos, cobrem o aviso quando o estado muda.'
Add-NiceTable @(
  @('#', 'Parte', 'Como testar sozinha'),
  @('1', 'Detectar que o estado mudou (exemplo: passou a "Apto para validacao").', 'O sistema (ou a fila) registra o novo estado no instante da mudanca, nao so "quando alguem olhar".'),
  @('2', 'Avisar a pessoa da etapa seguinte.', 'Essa pessoa recebe o recado sem abrir a planilha geral de manha.'),
  @('3', 'Ela ve o pedido certo, nao a planilha inteira.', 'Chega nome, matricula, obra e estado — o suficiente para validar ou liberar.')
) @(1.2, 7.5, 8.3) @(1)
Write-Mixed @(
  @{ t = 'Padrao:'; b = $true },
  @{ t = ' a informacao nao chega a quem precisa agir no momento certo.' }
)
Write-Mixed @(
  @{ t = 'Regra geral:'; b = $true },
  @{ t = ' sempre que uma solicitacao ficar pronta para a proxima etapa, uma versao visivel dela deve chegar automaticamente a quem precisa agir, sem depender de alguem abrir a planilha no horario de costume.' }
)

Write-H2 '7. Algoritmo (com decisao)'
Write-Plain '1. Supervisor envia a solicitacao (nome, matricula, funcao, obra de destino).'
Write-Plain '2. Assistente registra na planilha de mobilizacoes e marca "aguardando conferencia dos treinamentos".'
Write-Plain '3. Responsavel pelos treinamentos confere os 3 cursos obrigatorios daquela obra.'
Write-Mixed @(
  @{ t = '4. ' },
  @{ t = 'Se'; b = $true },
  @{ t = ' os 3 estao validos -> marcar "Apto para validacao" e avisar o coordenador de seguranca. ' },
  @{ t = 'Senao'; b = $true },
  @{ t = ' -> a solicitacao nao segue neste processo (a regularizacao fica fora).' }
)
Write-Plain '5. Coordenador valida a solicitacao.'
Write-Plain '6. O aviso chega a assistente. Assistente marca "Liberado para a obra" e confirma ao supervisor.'
Write-Plain 'A decisao obrigatoria esta no passo 4: so segue neste processo quem ja tem os 3 treinamentos validos. O aviso no mesmo passo e o que hoje nao existe — e e o gargalo.'

Write-Plain 'Fluxograma do algoritmo' 11 $true $script:Ink 6
Add-NiceTable @(
  @('Passo', 'O que acontece'),
  @('1', 'Supervisor envia a solicitacao'),
  @('2', 'Assistente registra e marca "aguardando conferencia"'),
  @('3', 'Treinamentos conferem os 3 cursos da obra'),
  @('4 — decisao', 'Os 3 treinamentos estao validos?')
) @(3.5, 13.5) @(1)

Write-Plain 'No passo 4 a resposta divide o fluxo:' 11 $false $script:Ink 6
Add-NiceTable @(
  @('NAO', 'SIM'),
  @('Nao segue neste processo', 'Marca "Apto" e avisa o coordenador')
) @(8.5, 8.5)

Add-NiceTable @(
  @('Passo', 'O que acontece (so se SIM)'),
  @('5', 'Coordenador valida'),
  @('6', 'Assistente marca "Liberado" e confirma ao supervisor')
) @(3.5, 13.5) @(1)

Write-H2 '8. Gatilho, condicao e acao'
Write-Plain 'No algoritmo, o pedaco que corta o gargalo e o fim do passo 4.'
Add-NiceTable @(
  @('Elemento', 'Neste processo', 'Na automacao'),
  @('Gatilho', 'A solicitacao acaba de ficar "Apto para validacao".', 'Nova linha na fila do coordenador (planilha).'),
  @('Condicao', 'Os 3 treinamentos daquela obra estao validos.', 'Filtro: estado = "Apto para validacao".'),
  @('Acao', 'Avisar o coordenador com uma versao visivel do pedido.', 'E-mail com nome, matricula, funcao, obra e estado.')
) @(3.2, 7.0, 6.8)
Write-Plain 'Nao se automatiza o processo inteiro. So o ponto em que hoje a informacao some: o estado muda na planilha e ninguem e avisado.'

Write-H2 '9. Automacao simples no Make (ligada ao gargalo)'
Write-Plain 'Ferramenta: Make. A empresa ja usa planilha e e-mail — a automacao liga as duas, em vez de inventar um sistema novo.'
Write-Mixed @(
  @{ t = 'Como passa a funcionar:'; b = $true },
  @{ t = ' quando os 3 treinamentos estao validos, a solicitacao entra numa fila visivel do coordenador (uma linha nova). O Make observa essa fila. Se o estado e "Apto para validacao", envia o e-mail ao coordenador na hora. No caso Rafael, essa espera foi 17 h 46 min. Com o aviso, o coordenador nao depende de abrir a planilha no dia seguinte de manha.' }
)
Add-NiceTable @(
  @('Modulo Make', 'Papel'),
  @('Google Sheets — Watch new rows (fila do coordenador)', 'Gatilho'),
  @('Filtro: Status igual a "Apto para validacao"', 'Condicao'),
  @('Gmail — Send an email (para o coordenador)', 'Acao')
) @(12.5, 4.5)
Write-Plain 'Prova de funcionamento: video curto (YouTube nao listado) em que uma linha nova entra na fila e o e-mail chega em seguida. Link do video no final deste documento e no formulario da trilha.'

Write-H2 'O que muda no caso Rafael'
Add-NiceTable @(
  @('Momento', 'Hoje (AS-IS)', 'Com o aviso (passo 4)'),
  @('Depois da conferencia 15:14', 'Planilha atualizada. Zero recado.', 'Fila + e-mail ao coordenador.'),
  @('Quando o coordenador ve', 'Manha seguinte, 09:00', 'Quando o estado muda, nao no horario de costume.'),
  @('Espera do gargalo', '17 h 46 min (67% do lead time)', 'Cai para o tempo de ele abrir o e-mail e validar (no caso, o trabalho ja era 5 min).')
) @(5.0, 6.0, 6.0)

Write-H2 'Fontes (so o briefing)'
Write-Plain '- Documento oficial do processo da Manutencao Guaras (caso Rafael Mendes, 5-6 de outubro).'
Write-Plain '- Horarios e estados copiados do relato. Nenhum tempo extra foi inventado.'
Write-Plain '- Entrevistas: gestor (nao ha aviso definido na mudanca de etapa); treinamentos ("quando termino, atualizo o status"); coordenador ("o problema e saber quando apareceu alguem novo").'

Write-Plain 'Link do video (YouTube nao listado)' 11 $true
Write-Plain 'Colar aqui depois de publicar: ________________________________'
Write-Plain 'Link das planilhas (Google Drive)' 11 $true
Write-Plain 'Colar aqui depois de publicar: ________________________________'
Write-Plain 'Cenario Make' 11 $true
Write-Plain 'Colar aqui depois de publicar: ________________________________'

Write-Plain 'Desafio 01 · Trilhas 2026 · Manutencao Guaras · processo de liberacao apos conferencia dos treinamentos obrigatorios. Documento unico da entrega (estudo de caso).' 9 $false $script:Mute 6

$savePath = [string]$out
$doc.SaveAs2($savePath, 16)
$doc.Close()
$word.Quit()
[System.Runtime.InteropServices.Marshal]::ReleaseComObject($sel) | Out-Null
[System.Runtime.InteropServices.Marshal]::ReleaseComObject($doc) | Out-Null
[System.Runtime.InteropServices.Marshal]::ReleaseComObject($word) | Out-Null
[GC]::Collect(); [GC]::WaitForPendingFinalizers()
Get-Process WINWORD -ErrorAction SilentlyContinue | Stop-Process -Force -ErrorAction SilentlyContinue

Get-Item -LiteralPath $out | Format-List FullName, Length, LastWriteTime
