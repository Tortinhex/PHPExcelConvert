# PHPExcelConvert
Conjunto de optimizações da API PHPExcel

DEFINIÇÃO
Esta ferramenta utiliza as funcionalidades da API PHPExcel (https://phpexcel.codeplex.com/), com funções extras, pra agilizar a conversão de HTML para .XLS



COMO IMPORTAR

Copie a pasta xlsexport para dentro da pasta do seu projeto. Na pagina HTML que contém a tabela, copie o seguinte código na seção de script:

function exportar() {
    
    //Copia o HTML da tabela para variavel
    var table = document.getElementById('tabela').outerHTML;
    //Retira qualquer espaco e quebra de linha
    table = table.replace(/[\r\t\n]|  /g, '');
    
    //Cria um form na pagina
    var frag = document.createDocumentFragment();
    var form = document.createElement("form");
    form.name="convertXls";
    form.method="post";
                
    var input = document.createElement("input");
    input.type="hidden";
    input.name="html";
                
    form.appendChild(input);
    frag.appendChild(form);
    document.body.appendChild(frag);

    //Submit
    document.convertXls.html.value = table;
    document.convertXls.action = "xlsexport/Conversor.php";
    document.convertXls.submit();
}

OBS: O método acima deve ser chamado por um botao:
  <button type="button" onclick="exportar();">Click</button>



COMO CONFIGURAR A API

Setar o nome do documento (nome_do_documento.xls):
$conversor->setNameDocument("nome_do_documento");

Setar o nome da aba de trabalho:
$conversor->setNameAba("Pasta de trabalho");

Setar a cor do fundo:
$conversor->setColorFundo("FFFFFF");

Setar a cor de fundo do titulo:
$conversor->setColorHeader("AAAAAA");

Setar a cor do striped:
$conversor->setColorsStrip("FFFFFF", "EEEEEE");

Setar o tamanho da fonte do titulo:
$conversor->setFontSizeTitle(16);

Setar o tamanho da fonte do corpo:
$conversor->setFontSizeBody(12);

Setar a fonte do codigo:
$conversor->setFontFamily("verdana");

Setar a posicao inicial em X:
$conversor->setDefaultPositionX(3);

Setar a posicao inicial em Y:
$conversor->setDefaultPositionY(3);

Gera o arquivo XLS:
$conversor->gerarXls();
