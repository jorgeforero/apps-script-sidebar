/**
 * SIDEBAR
 */

/**
* onOpen - Custom menu
**/
function onOpen() {
  let ui = SpreadsheetApp.getUi();
  ui.createMenu('Acciones')
  .addItem('🏃🏻 Ver Info Atletas', 'showSB_info')
  .addToUi();
};

/**
* showSB_info - Información de atletas y envios
**/
function showSB_info() {
  // Se calcula la información y se genera el template
  let actualData = getData();
  let template = getTplGraph(actualData)
  let ui = HtmlService.createTemplate(template)
                      .evaluate()
                      .setSandboxMode(HtmlService.SandboxMode.IFRAME)
                      .setTitle('Información');
  SpreadsheetApp.getUi().showSidebar(ui);
};

/**
* getData()
* Obtiene los conteos de: Total de atletas, Total atletas activos, 
* Total de correos enviados día actual, y la quota de correos restantes
* 
* @param {void} - void
* @return {object} - Objeto con los datos calculados
*   athTotal: Total atletas
*   athActives: Total atletas activos
*   emailsSent: Total envíos de hoy
*   leftEmails: Quota restante
**/
function getData() {
  let athletes  = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet().getDataRange().getValues();
  let header = athletes[ 0 ];
  let todayFormatted = Utilities.formatDate( new Date(), 'GMT-5', 'MM/dd/yyy' );
  let counterTotal = 0;
  let counterActives = 0;
  let counterSentToday = 0;
  for ( let indx=1; indx< athletes.length; indx++ ) {
    let athlete = getRowAsObject( athletes[indx], header );
    // Conteo de los atletas activos
    if ( athlete.active == true ) counterActives++;
    // Conteo envíos fecha actual
    if ( Utilities.formatDate(new Date( athlete.lastemail ), 'GMT-5', 'MM/dd/yyy') == todayFormatted) counterSentToday++;
    // Conteo total atletas
    counterTotal++;
  };//for
  return { total: counterTotal,
           actives: counterActives,
           sent: counterSentToday,
           left: MailApp.getRemainingDailyQuota()};
};

/**
* getGraphs
* Recibe los datos generados y los reemplaza en el template.  Retorna una cadena HTML con los valores
* reemplazados
*
* @param {object} - Objeto con los datos a reemplazar
*   athTotal: Total atletas
*   athActives: Total atletas activos
*   emailsSent: Total envíos de hoy
*   leftEmails: Quota restante
* @return {string} - Cadena HTML con los datos generados
**/
function getTplGraph(data) {
  let results = [ [ 'Total envíos', data.sent, data.actives, data.actives * 0.3, data.actives * 0.6 ],
                  [ 'Quota restantes', data.left, 1500, 5, 50 ] ];  
  // Template de los resultados
  let tpl_row = HtmlService.createHtmlOutputFromFile( 'tpl_sb_row.html' ).getContent(); 
  let listOut = '';
  for ( let indx=0; indx<results.length; indx++ ) {
    let rowHtml = tpl_row;
    let row = results[ indx ];
    let calPercent = ( row[1] *100 ) / row[2];
    // Reemplazables
    rowHtml = rowHtml.replace( '##ITEM##', row[0] );
    rowHtml = rowHtml.replace( '##VALUE##', row[1] );
    rowHtml = rowHtml.replace( '##COLOR##', getStatusColor( row[1], row[3], row[4] ) );
    rowHtml = rowHtml.replace( '##PERCENT##', numberFormat( calPercent ) );
    listOut += rowHtml;
  };
  // Template final
  let tpl_info = HtmlService.createHtmlOutputFromFile('tpl_sb_info.html').getContent();
  tpl_info = tpl_info.replace( '##SENTTODAY##', listOut );
  tpl_info = tpl_info.replace( '##ATHTOTAL##', data.total );
  tpl_info = tpl_info.replace( '##ATHACTIVE##', data.actives );
  tpl_info = tpl_info.replace( '##ATHINACTIVE##', data.total - data.actives );
  return tpl_info;
}

/**
* getStatusColor
* Dado el valor del resultado y los limites, retorna un color representativo del resultado (tipo semaforo)
* red < TrsA <= yellow <= TrsB < green
*
* @param {number} Value - Valor a evualuar
* @param {number} TrsA - Limite A 
* @param {number} TrsB - Limite B
* @return {string} - Color obtenido - SEMAFORO
**/
function getStatusColor( Value, TrsA, TrsB ) {
  let color;
  // Traffic lights
  if ( Value < TrsA) color = 'red';
  if ( ( Value >= TrsA ) && ( Value <= TrsB ) ) color = 'yellow';
  if ( Value > TrsB ) color = 'green';
  return color;
};

/**
* numberFormat
* Formatea el número dado para que tenga la apariencia de moneda
*
* @param {number} num - Número a formatear
* @return {strting} - Cadena que representa el valor en formato moneda
**/
function numberFormat(num) {
  if (num == 0) return '0.0';
  else return num.toFixed(1).replace(/(\d)(?=(\d{3})+(?!\d))/g, '$1,');
};
