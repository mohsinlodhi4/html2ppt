<?php

require __DIR__.'/vendor/autoload.php';

\PhpOffice\PhpPresentation\Autoloader::register();
\PhpOffice\Common\Autoloader::register();

use PhpOffice\PhpPresentation\PhpPresentation;
use PhpOffice\PhpPresentation\IOFactory;
use PhpOffice\PhpPresentation\Style\Color;
use PhpOffice\PhpPresentation\Style\Border;
use PhpOffice\PhpPresentation\Style\Alignment;
use PhpOffice\PhpPresentation\Style\Fill;

// echo "<pre>";
// print_r( get_class_methods(PhpPresentation::class) );

$data = [
    [ // slide1
        ['logo'=>  './logo.png'],
        ['h1'=> 'Parabéns, sua avaliação está pronta! '],
        ['h2' => ' Perfil'],
        ['p2' => ' Empresa Teste Ltda - Santa Cruz do Sul, RS
Fundada em 2020'],
        ['h2'=>'Operações'],
        ['p3'=>['text'=>'Estágio da empresa: Crescimento inicial 
Número de sócios: 2 
Experiência da equipe: Mais de 10 anos
Propriedades intelectuais: Marca registrada, Certificações, Domínio / Website / Softwares', 'lines'=>4]],
        ['h2'=>'Oportunidade'],
        ['p4'=>'Média concorrência
Negócio não escalável'],
        ['table'=> [
                ['SETOR', 'EMPRESAS', 'D/E'],
                ['Software (Internet)', '40', '15.22%'],
                ['Software (Sistemas e Aplicações)', '462', '4.43%'],
        ]],
        ['center'=>  'Beta dos setores: 1.28 β']
    ],

    [ // slide2
        ['h1'=> 'Desempenho '],
        ['p1' => ['lines'=>1,'text'=>'Receita                                                  297.000,00 R$']],
        ['p2' => ['lines'=>1,'text'=>'Dívidas                                                  15.000,00 R$']],
        ['p3' => ['lines'=>1,'text'=>'Dinheiro em caixa                                        20.000,00 R$']],
        ['p4' => ['lines'=>1,'text'=>'Retorno Sobre o Capital Investido (ROIC)                 80,00 R$']],
        ['p5' => ['lines'=>1,'text'=>'EBIT                                                     84.800,00 R$']],
        ['p6' =>      ['lines'=>1,'text'=>'Margem EBIT                                              28.55%']],
        [ 'table'=> [
            [ ['Taxa Livre de Risco', '12.56%'], ['Prêmio de Risco de Capital Próprio', '7.21%'], ['Desconto de Iliquidez', '15%'] ],
            [ ['Risco País', '2.97%'], ['Custo do Capital', '19,28%'], ['Custo da Dívida', '15,88%'] ],
            [ ['Custo do Capital Próprio', '19,75%'], ['IPCA acumulado 12 meses', '11.30%'], ['Cotação diária do dólar', 'R$ 5,13'] ],
        ] ],
    ],
    
    [ // slide3
        [ 'h1'=> '1. Valor econômico com DCF' ],
        [ 'p1'=> ['text'=>'No DCF (Discounted Cash Flow, ou Fluxo de Caixa Descontado), o objetivo é identificar o valor de uma empresa considerando seu fluxo de caixa, crescimento e risco. Ou seja, estima o seu potencial econômico em gerar retorno para sócios e credores. Esse critério engloba fatores como plano de negócio, competência da gestão, tradição, marca, carteira de clientes e qualquer decisão que impacte na geração de resultados.', 'lines'=>6] ],
        [ 'p1'=>['text'=>'Conforme Aswath Damodaran, uma boa avaliação é o resultado de dois elementos: o desempenho histórico da empresa e a criação de uma narrativa sólida (através das projeções financeiras).', 'lines'=>3] ],
        ['centerBox'=> ['text'=>'Valor calculado com Fluxo de Caixa Descontado:', 'text2'=> 'R$ 448.216,05', 'text2Font'=>20, 'bg'=>'FFd7f5fc', 'text2Color'=>'FF03c3ec', 'width'=>600, 'height'=>90 ]],
        ['image' => ['height'=>450, 'src'=>'./chart1.png', 'width'=>940 ] ],
    ],
    [ //slide 4
        [ 'h1'=> '2. Valor de mercado com Múltiplos' ],
        [ 'p1'=> ['text'=>'Um dos componentes mais desafiadores do valuation de pequenas e médias empresas, é que não existem empresas exatamente iguais; Dessa forma, é necessário encontrar outras empresas semelhantes para poder realizar a comparação. Mais especificamente, usando o EV/Sales (múltiplo de faturamento bastante utilizado inclusive no mercado de ações), que agrega ao valor da empresa os reflexos de dívidas e dinheiro em caixa.', 'lines'=>5] ],
        [ 'p1'=>['text'=>'O Múltiplo de Valor de Mercado (EV/Sales) mais adequado calculado é: 2.42x (variável conforme perfil da empresa).', 'lines'=>2] ],
        [ 'p1'=>['text'=>'Usando uma amostra de dezenas de pequenas e médias empresas previamente avaliadas, a Valutech calcula o múltiplo da empresa de acordo com o faturamento, setor e geolocalização.', 'lines'=>2] ],
        ['centerBox'=> ['text'=>'Valor calculado com Múltiplos de Mercado:', 'text2'=> 'R$ 718.740,00', 'text2Font'=>20, 'bg'=>'FFd7f5fc', 'text2Color'=>'FF03c3ec', 'width'=>600, 'height'=>90 ]],
    ],
    [ //slide 5
        [ 'h1'=> 'Média da avaliação da empresa' ],
        [ 'p1'=> ['text'=>'Para se chegar ao resultado médio do valor da empresa, foram utilizados os métodos de Fluxo de Caixa Descontado (DCF), baseado no resultado histórico e expectativas de ganhos futuros fornecidos pelo usuário, e o método de Múltiplos de Mercado, que se baseia na forma como outras empresas similares são precificadas (este relatório não leva em consideração eventos futuros e seus impactos após sua data de emissão).', 'lines'=>5] ],
        [ 'p1'=>['text'=>'Cada método tem um peso igual no valor final, ou seja:', 'lines'=>1] ],
        ['double'=> [
            ['image' => ['height'=>100, 'src'=>'./chart_mini.png', 'width'=> 80] ],
            ['pMulti'=> ['text'=>'50% representa o potencial econômico da empresa (DCF); e', 'text2'=> '50% representa seu valor teórico de mercado (Múltiplos).'] ]
        ]],
        ['centerBox'=> ['addYOffset'=>25, 'text'=>'Valor médio calculado da empresa Empresa Teste Ltda:', 'text2'=> 'R$ 583.478,03', 'text2Font'=>20, 'bg'=>'FFe8fadf', 'text2Color'=>'FF71dd37', 'text3'=>'(Valor considerado na data de emissão deste documento e sujeito às premissas aqui descritas)', 'text3Font'=>12, 'width'=>700, 'height'=>110 ]],

        // [ 'p1'=>['text'=>'Usando uma amostra de dezenas de pequenas e médias empresas previamente avaliadas, a Valutech calcula o múltiplo da empresa de acordo com o faturamento, setor e geolocalização.', 'lines'=>2] ],
        // ['centerBox'=> ['text'=>'Valor calculado com Múltiplos de Mercado:', 'text2'=> 'R$ 718.740,00', 'text2Font'=>20, 'bg'=>'FFd7f5fc', 'text2Color'=>'FF03c3ec', 'width'=>600, 'height'=>90 ]],
    ],
    [ // slide3
        [
            'doubleSlide'=>[
                [ //inner slide 1
                    'data'=>[
                        [ 'h1'=> 'Avaliação da Marca' ],
                        [ 'p1'=> ['text'=>'A metodologia desenvolvida pela Valutech é uma variação do método Company Valuation Less Value Of Net Tangible Assets (ou Avaliação da Empresa Menos Valor dos Ativos Tangíveis Líquidos), descrita no livro "The International Brand Valuation Manual" (Salinas, 2009), o qual foi adaptado especialmente para micro, pequenas e médias empresas.', 'lines'=>4] ],
                        [ 'p1'=>['text'=>'Resumidamente, o valor da marca é estabelecido determinando o valor da empresa e então subtraindo o valor líquido dos ativos empregados pela empresa em sua operação.', 'lines'=>3] ],
                        [ 'p1'=>['text'=>'A Valutech então simplificou o processo de subtração de ativos empregados: ao final do processo de valuation do negócio, iremos subtrair itens comuns como os ativos imobilizados, ou seja, imóveis (salas, pavilhões, prédios, fábricas), maquinário, veículos, equipamentos e mobília de escritório, computadores, material de escritório e ferramentas.', 'lines'=>4] ],
                        ['center'=>  [ 'addYOffset'=> -10, 'text'=>'Valor calculado da Marca:', 'text2'=>'R$ 418.216,05', 'text2Font'=>20, 'text2Color'=>'FF03c3ec' ]],
                    ],
                    'settings'=>[
                        'width'=>650
                    ]

                ],
                // [ //inner slide 
                //     'data'=>[
                //         ['image' => ['height'=>400, 'src'=>'./chart2.png', 'width'=>480 ] ],
                //     ]
                // ],
            ]
        ],
        ['image' => ['height'=>400, 'src'=>'./chart2.png', 'width'=>430, 'addYOffset'=>180, 'addXOffset'=>520, ] ],
    ],
    [ // slide 4
        [ 'h1'=> 'Complementos' ],
        [ 'h1'=> 'Tabela de valor:' ],
        ['table'=> [
            ['PERÍODO', 'ANO-BASE', '2023', '2024', '2025', '2026', '2027', 'PERPETUIDADE (VT)'],
            ['RECEITAS (EM R$)', '297.000,00', '330.561,00', '367.914,39', '409.488,72', '455.760,95', '507.261,94', '537.697,66'],
            ['CRESC. RECEITA', '11.30%', '11.30%', '11.30%', '11.30%', '11.30%', '11.30%', '6%'],
            ['EBIT', '84.800,00', '89.040,00', '93.492,00', '98.166,60', '103.074,93', '108.228,68', '113.640,11'],
            ['EBIT %', '5.00%', '5.00%', '5.00%', '5.00%', '5.00%', '5.00%', '5.00%'],
            ['ALÍQUOTA IMPOSTOS', '6.00%', '6.00%', '6.00%', '6.00%', '6.00%', '6.00%', '6.00%'],
            ['EBIT (1-T) NOPAT', '79.712,00', '83.697,60', '87.882,48', '92.276,60', '96.890,43', '101.734,96', '106.821,71'],
            ['(-) REINVESTIMENTOS %', '-', '15.5%', '14.8%', '14.1%', '13.4%', '12.8%', '12.8%'],
            ['FCFF (FCL)', '-', '70.724,47', '74.875,87', '79.265,60', '83.907,12', '88.712,88', '93.148,53'],
            ['VALOR TERMINAL *', '', '', '', '', '', '', 'R$ 701.419,64'],
            ['CUSTO DO CAPITAL', '-', '19.28%', '19.28%', '19.28%', '19.28%', '19.28%', '19.28%'],
            ['FATOR DE DESCONTO ACUMULADO', '-', '0.838', '0.703', '0.589', '0.494', '0.414', ''],
            ]],
        ],
    [ // slide 5
        ['table'=>[
            ['VALOR PRESENTE (FCFF)', '-', '59.292,82', '52.626,76', '46.706,99', '41.450,36', '36.740,79', '290.495,74'],
            ['VALOR DA EMPRESA', '', '', '', '', '', '', 'R$ 527.313,00'],
            ['DESCONTO DE ILIQUIDEZ', '', '', '', '', '', '', '- 15%'],
            ['VALOR FINAL DA EMPRESA COM DCF', '', '', '', '', '', '', 'R$ 448.216,05'],
        ]],
    ],
    [ // slide 6
        [ 'h2'=> 'Cenário de crescimento escolhido:' ],
        [ 'p1'=> ['text'=>'Crescimento acompanhando a inflação (11.30% / 5.00% EBIT).', 'lines'=>1] ],
        [ 'h1'=> 'Taxa Livre de Risco - Brasil 5 anos:' ],
        ['image' => ['height'=>450, 'src'=>'./chart3.png', 'width'=>540 ] ],
        [ 'h1'=> ['text'=> 'Referências bibliográficas:', 'addYOffset'=>-150] ],
        [ 'p1'=>['text'=>'Damodaran, A (2007). Avaliação de Empresas - 2ª Edição. São Paulo: Pearson Prentice Hall;', 'lines'=>1] ],
        [ 'p1'=>['text'=>'Damodaran, A. (2011). The Little Book of Valuation: How to Value a Company, Pick a Stock and Profit. John Wiley & Sons, Inc.;', 'lines'=>2] ],
        [ 'p1'=>['text'=>'Marques, K. C. (2015). Análise Financeira das Empresas 2ª Edição. São Paulo: Freitas Bastos Editora;', 'lines'=>1] ],
        [ 'p1'=>['text'=>'Damodaran, A. (2017). Narrative and Numbers: The Value of Stories in Business. New York: Columbia University Press;', 'lines'=>2] ],
        [ 'p1'=>['text'=>'Neto, A. A. (2019). Valuation: Métricas & Avaliação de Empresas. 2ª Edição. São Paulo: Editora Atlas;', 'lines'=>1] ],
        [ 'p1'=>['text'=>'Salinas, G. (2009). The International Brand Valuation Manual. Madrid: Editora Wiley;', 'lines'=>1] ],
        
    ],
    [ // slide 7
        ['h1'=> 'Glossário de termos:'],
        [ 'h3'=> ['text'=> 'Ativos Operacionais', 'color'=>'FF66a7f'] ],
        [ 'p1'=>['text'=>'Ativos necessários para operar o negócio fundamental da empresa;', 'lines'=>1] ],
        [ 'h3'=> ['text'=> 'Desconto de Iliquidez ', 'color'=>'FF66a7f'] ],
        [ 'p1'=>['text'=>'Desconto aplicado em ativos que não podem ser facilmente vendidos. Por isso, acrescenta-se esta redução para tornar a venda do ativo mais fácil;', 'lines'=>2] ],
        [ 'h3'=> ['text'=> 'DRE', 'color'=>'FF66a7f'] ],
        [ 'p1'=>['text'=>'Demonstração do Resultado do Exercício;', 'lines'=>1] ],
        [ 'h3'=> ['text'=> 'EBIT', 'color'=>'FF66a7f'] ],
        [ 'p1'=>['text'=>'Earnings Before Interest and Taxes (Lucro antes do Juros e dos Impostos – LAJIR);', 'lines'=>1] ],
        [ 'h3'=> ['text'=> 'EV/Sales', 'color'=>'FF66a7f'] ],
        [ 'p1'=>['text'=>'Múltiplo de faturamento. Demonstra a relação do valor da empresa (Enterprise Value) com seu faturamento anual;', 'lines'=>2] ],
        [ 'h3'=> ['text'=> 'FCFF', 'color'=>'FF66a7f'] ],
        [ 'p1'=>['text'=>'Free Cash Flow to the Firm (Fluxo de Caixa Livre para a Empresa - FCLE);', 'lines'=>1] ],
        [ 'h3'=> ['text'=> 'IPCA', 'color'=>'FF66a7f'] ],
        [ 'p1'=>['text'=>' Índice de Preços ao Consumidor Amplo (Índice de inflação calculado pelo IBGE);', 'lines'=>1] ],
        [ 'h3'=> ['text'=> 'NOPAT ', 'color'=>'FF66a7f'] ],
        [ 'p1'=>['text'=>'  Net Operating Profit After Taxes (Lucro Operacional Líquido Após Impostos);', 'lines'=>1] ],

    ],
    [ // slide 8
        [ 'h3'=> ['text'=> 'PIB ', 'color'=>'FF66a7f'] ],
        [ 'p1'=>['text'=>' Produto Interno Bruto;', 'lines'=>1] ],
        [ 'h3'=> ['text'=> 'Reinvestimento', 'color'=>'FF66a7f'] ],
        [ 'p1'=>['text'=>'nvestimentos necessários à continuidade, manutenção e crescimento da capacidade produtiva da empresa;', 'lines'=>2] ],
        [ 'h3'=> ['text'=> 'ROA ', 'color'=>'FF66a7f'] ],
        [ 'p1'=>['text'=>' Return on Assets (Retorno sobre os ativos);', 'lines'=>1] ],
        [ 'h3'=> ['text'=> 'ROE', 'color'=>'FF66a7f'] ],
        [ 'p1'=>['text'=>' Return on Equity (Retorno sobre o Patrimônio Líquido);', 'lines'=>1] ],
        [ 'h3'=> ['text'=> 'ROIC', 'color'=>'FF66a7f'] ],
        [ 'p1'=>['text'=>' Return on Invested Capital (Retorno sobre o Capital Investido);', 'lines'=>1] ],
        [ 'h3'=> ['text'=> 'Sales to Capital Ratio', 'color'=>'FF66a7f'] ],
        [ 'p1'=>['text'=>'Taxa de vendas (receitas) em comparação ao capital investido. Um indicador utilizado para auferir o retorno do capital investido', 'lines'=>2] ],
        [ 'h3'=> ['text'=> 'Valor Terminal', 'color'=>'FF66a7f'] ],
        [ 'p1'=>['text'=>'Valor final na avaliação de fluxo de caixa descontado, uma vez que é impossível estimar fluxos de caixa para sempre.', 'lines'=>2] ],

    ],

];


createPpt($data);



function createPpt($slides, $title='slide'){
    $objPHPPowerPoint = new PhpPresentation();

    $i=0;
    foreach($slides as $slide){
        $currentSlide = $i==0 ? $objPHPPowerPoint->getActiveSlide() : $objPHPPowerPoint->createSlide();

        $currentYOffset = 0;        
        foreach($slide as $element){
            $tag = array_key_first($element);
            $value = $element[$tag];

            if($tag=='logo'){
                $currentYOffset += 10; 
                $height = 36;
                $shape = $currentSlide->createDrawingShape();
                $shape->setName('PHPPresentation logo')
                    ->setDescription('PHPPresentation logo')
                    ->setPath($value)
                    ->setHeight(36)
                    ->setOffsetX(400)
                    ->setOffsetY($currentYOffset);
            }
            if($tag=='image'){
                $currentYOffset += 10; 
                if(isset($value['addYOffset'])){
                    $currentYOffset += $value['addYOffset']; 
                }
                $currentXOffset = 10;
                if(isset($value['addXOffset'])){
                    $currentXOffset = $value['addXOffset']; 
                }

                $height = $value['height'];
                createImageTag($currentSlide, $tag, $value, $height, $currentYOffset, $currentXOffset);
                // $shape = $currentSlide->createDrawingShape();
                // $shape->setName('PHPPresentation logo')
                //     ->setDescription('PHPPresentation logo')
                //     ->setPath($value['src'])
                //     ->setHeight($height+50)
                //     ->setWidth(940)
                //     ->setOffsetX(10)
                //     ->setOffsetY($currentYOffset);
            }


            if( in_array($tag ,['h1','h2','h3','h4']) ){
                $currentYOffset += 5; 
                $height = 40;
                $text = '';
                if(is_array($value) && isset($value['addYOffset'])){
                    $currentYOffset += $value['addYOffset']; 
                }
                $text = $value['text'] ?? $value;
                $color = $value['color'] ?? 'FF47008a';

                createHeading($currentSlide,$tag, $text, $height, $currentYOffset, 10, $color);
            }

            if( in_array($tag ,['p1','p2','p2_line','p3','p4','p5','p6', 'center']) ){
                $currentYOffset += 5; 
                if(isset($value['addYOffset'])){
                    $currentYOffset += $value['addYOffset']; 
                }
                $height = 40;                
                createPara($currentSlide,$tag, $value, $height, $currentYOffset);
            }

            if($tag=='centerBox'){
                $currentYOffset += 5; 
                if(isset($value['addYOffset'])){
                    $currentYOffset += $value['addYOffset']; 
                }
                $height = $value['height'] ?? 40;
                $width = $value['width'];
                $text = $value['text'];
                
                $fontSize = 14; 
                $shape = $currentSlide->createRichTextShape()
                ->setHeight($height)
                ->setWidth($width)
                ->setOffsetX(150)
                ->setOffsetY($currentYOffset);
                //line 1
                $shape->getActiveParagraph()->getAlignment()->setHorizontal( Alignment::HORIZONTAL_CENTER );
                $textRun = $shape->createTextRun($text);
                $textRun->getFont()
                ->setSize($fontSize)
                ->setColor( new Color( 'FF8592a3' ) );
                // line 2
                if(isset($value['text2'])){
                    $shape->createParagraph()->setLineSpacing(120)->getAlignment()->setHorizontal( Alignment::HORIZONTAL_CENTER );
                    $textRun = $shape->createTextRun($value['text2']);
                    $textRun->getFont()->setBold(true)
                    ->setSize($value['text2Font'])
                    ->setColor( new Color( $value['text2Color'] ?? 'FF03c3ec' ) );
                }
                if(isset($value['text3'])){
                    $shape->createParagraph()->setLineSpacing(120)->getAlignment()->setHorizontal( Alignment::HORIZONTAL_CENTER );
                    $textRun = $shape->createTextRun($value['text3']);
                    $textRun->getFont()->setBold(true)
                    ->setSize($value['text3Font'] ?? 12)
                    ->setColor( new Color( $value['text3Color'] ?? 'FF8592a3' ) );
                }

                if(isset($value['bg'])){
                    $bgColor = new Color($value['bg']);
                    $shape->getFill()->setFillType(Fill::FILL_SOLID)
                        ->setStartColor($bgColor)
                        ->setEndColor($bgColor);
                }
            }

            if($tag == 'table'){
                $currentYOffset += 15;
                $height = sizeof($value)*40;
                if( is_array($value[0][0]) ){
                    $height = sizeof($value)*50;
                }
                $columns = $value[0];
                $shape = $currentSlide->createTableShape( sizeof($columns) );
                $shape->setOffsetY($currentYOffset)
                ->setOffsetX(10)
                ->setWidth(800);

                createTable($shape, $value);
            }

            if($tag=='double'){
                $currentYOffset += 10; 
                $height = 100;

                $currentXOffset=10;
                foreach($value as $element){
                    $innerTag = array_key_first($element);
                    $innerValue = $element[$innerTag];
                    
                    if($innerTag=='image') createImageTag($currentSlide, $innerTag, $innerValue, $height, $currentYOffset, $currentXOffset);
                    else if ($innerTag=='pMulti' || $innerTag=='p1') createPara($currentSlide, $innerTag, $innerValue, $height, $currentYOffset, $currentXOffset);
                    $currentXOffset += $innerValue['width'] ?? 0;
                }
            }
            if($tag == 'doubleSlide'){
                $currentYOffset += 10; 
                $currentXOffset=10;
                foreach($value as $innerSlide){
                    $data = $innerSlide['data'];
                    $settings = $innerSlide['settings'] ?? null;
                    createInnerSlide($currentSlide, $data, $currentYOffset, $currentXOffset);
                    
                    $currentXOffset += $settings['width'] ?? 0;
                }

            }

            $currentYOffset += $height;
        }

        $i++;
    }
    
    
    $oWriterPPTX = IOFactory::createWriter($objPHPPowerPoint, 'PowerPoint2007');
    $oWriterPPTX->save(__DIR__ . "/$title.pptx");
    
}

function createTable($shape, $table){
    $j=0;

    $isTwoDimenional = is_array($table[0][0]);
    foreach($table as $rowData){
        $row = $shape->createRow();
        foreach($rowData as $cell){
            $cellValue = $isTwoDimenional ? $cell[0] : $cell;
            $secondLine = $isTwoDimenional ? $cell[1] : '';
            $colorHex = $j==0 && !$isTwoDimenional ? 'FF47008a' : 'FF000000';

            $cell = $row->nextCell();
            $textRun = $cell->createTextRun($cellValue);
            $textRun->getFont()->setBold(true);
            $textRun->getFont()->setSize(12);
            $textRun->getFont()->setColor(new Color( $colorHex ));
            if($isTwoDimenional){
                $cell->createBreak();
                $textRun = $cell->createTextRun($secondLine);
                $textRun->getFont()->setBold(true);
                $textRun->getFont()->setSize(14);
                $textRun->getFont()->setColor(new Color('FF47008a'));
            }

            $cell->getActiveParagraph()->getAlignment()
            ->setMarginLeft(10)->setMarginBottom(5);
        }
        $j++;
    }
}

function createPara($currentSlide, $tag, $value, &$height, $currentYOffset, $currentXOffset=10){
    if(is_array($value)){
        $height = 25*($value['lines'] ?? 1);
        $text = $value['text'];
    }else{
        $text = $value;
    }
    $fontSize = $value['fontSize'] ?? 14;
    $shape = $currentSlide->createRichTextShape()
                ->setHeight($height)
                ->setWidth(800)
                ->setOffsetX($currentXOffset)
                ->setOffsetY($currentYOffset);
                $shape->getActiveParagraph()->getAlignment()->setHorizontal( $tag=='center' ? Alignment::HORIZONTAL_CENTER : Alignment::HORIZONTAL_LEFT  );
                $textRun = $shape->createTextRun($text);
                $textRun->getFont()
                ->setSize($fontSize)
                ->setColor( new Color( $value['color'] ?? 'FF000000' ) );

                if(isset($value['text2'])){
                    $shape->createParagraph()->getAlignment()->setHorizontal( $tag=='center' ? Alignment::HORIZONTAL_CENTER : Alignment::HORIZONTAL_LEFT  );
                    $textRun = $shape->createTextRun($value['text2']);
                    $textRun->getFont()
                    ->setSize( $value['text2Font'] ?? $fontSize)
                    ->setColor( new Color( $value['text2Color'] ?? 'FF000000' ) );
                }
}

function createImageTag($currentSlide, $tag, $value, &$height, $currentYOffset, $currentXOffset=10){
    $shape = $currentSlide->createDrawingShape();
    $shape->setName('PHPPresentation logo')
        ->setDescription('PHPPresentation logo')
        ->setPath($value['src'])
        ->setHeight($height)
        ->setWidth($value['width'] ?? 900)
        ->setOffsetX($currentXOffset)
        ->setOffsetY($currentYOffset);
}
function createHeading($currentSlide, $tag, $value, &$height, $currentYOffset, $currentXOffset=10, $color ='FF47008a' ){
    $fontSize = $tag=='h1' ? 20 :  18; 
    $fontSize = $tag=='h2' ? $fontSize : 16;
                $shape = $currentSlide->createRichTextShape()
                ->setHeight($height)
                ->setWidth(600)
                ->setOffsetX($currentXOffset)
                ->setOffsetY($currentYOffset);
                $shape->getActiveParagraph()->getAlignment()->setHorizontal( Alignment::HORIZONTAL_LEFT );
                $textRun = $shape->createTextRun($value);
                $textRun->getFont()->setBold(true)
                ->setSize($fontSize)
                ->setColor( new Color( $color ) );
}

function createInnerSlide($currentSlide, $slidesData, $currentYOffset, $currentXOffset=10){

    foreach($slidesData as $element){
        $innerTag = array_key_first($element);
        $innerValue = $element[$innerTag];
        $height =$innerValue['height'] ?? 30;

        if($innerTag=='h1') createHeading($currentSlide, $innerTag, $innerValue, $height, $currentYOffset, $currentXOffset);
        if($innerTag=='p1' || $innerTag=='center' ) createPara($currentSlide, $innerTag, $innerValue, $height, $currentYOffset, $currentXOffset);
        $currentYOffset += $height;
    }

}
