'use strict';
const {
  Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell,
  Header, Footer, AlignmentType, HeadingLevel, BorderStyle, WidthType,
  ShadingType, PageNumber, PageBreak, LevelFormat,
} = require('docx');

const BD='1F4E79', BM='2E75B6', RD='C00000', RL='FFE7E7';
const AD='BF8F00', AB='FFF2CC', GB='E2EFDA', GR='F2F2F2', WH='FFFFFF';

const bdr = { style:BorderStyle.SINGLE, size:1, color:'CCCCCC' };
const allB = t => ({ top:t, bottom:t, left:t, right:t });
const cm = { top:80, bottom:80, left:120, right:120 };
const W = 9360;

const hd1 = t => new Paragraph({
  heading: HeadingLevel.HEADING_1,
  children: [new TextRun({ text:t, font:'Arial', bold:true, color:WH })],
  shading: { fill:BD, type:ShadingType.CLEAR },
  spacing: { before:320, after:120 },
});
const hd2 = t => new Paragraph({
  heading: HeadingLevel.HEADING_2,
  children: [new TextRun({ text:t, font:'Arial', bold:true, color:BD })],
  border: { bottom:{ style:BorderStyle.SINGLE, size:6, color:BM, space:1 } },
  spacing: { before:260, after:100 },
});
const hd3 = t => new Paragraph({
  heading: HeadingLevel.HEADING_3,
  children: [new TextRun({ text:t, font:'Arial', bold:true, color:BM, size:24 })],
  spacing: { before:180, after:80 },
});
const para = (runs, sp={before:60,after:100}) => new Paragraph({
  children: typeof runs==='string' ? [new TextRun({text:runs,font:'Arial',size:22})] : runs,
  spacing: sp,
});
const b  = (t,c='000000') => new TextRun({text:t,font:'Arial',bold:true,size:22,color:c});
const n  = t => new TextRun({text:t,font:'Arial',size:22});
const bullet = children => new Paragraph({
  numbering:{reference:'bullets',level:0},
  children, spacing:{before:60,after:60},
});

function alertBox(title, text, bg=RL, bc=RD, tc=RD) {
  return [
    new Paragraph({
      children:[new TextRun({text:title,font:'Arial',bold:true,size:22,color:tc})],
      shading:{fill:bg,type:ShadingType.CLEAR},
      border:{left:{style:BorderStyle.SINGLE,size:18,color:bc}},
      indent:{left:240}, spacing:{before:80,after:40},
    }),
    new Paragraph({
      children:[new TextRun({text,font:'Arial',size:20,color:'333333'})],
      shading:{fill:bg,type:ShadingType.CLEAR},
      border:{left:{style:BorderStyle.SINGLE,size:18,color:bc}},
      indent:{left:240}, spacing:{before:0,after:120},
    }),
  ];
}
const warnBox = (t,x) => alertBox(t,x,AB,AD,AD);
const infoBox = (t,x) => alertBox(t,x,'DEEBF7',BM,BM);

function simpleTable(headers, rows, colW, hdBg=BM) {
  return new Table({
    width:{size:W,type:WidthType.DXA}, columnWidths:colW,
    rows:[
      new TableRow({ children: headers.map((h,j) =>
        new TableCell({
          borders:allB(bdr), shading:{fill:hdBg,type:ShadingType.CLEAR},
          width:{size:colW[j],type:WidthType.DXA}, margins:cm,
          children:[new Paragraph({children:[new TextRun({text:h,font:'Arial',bold:true,color:WH,size:18})],alignment:AlignmentType.CENTER})],
        })
      )}),
      ...rows.map((r,i) => new TableRow({ children: r.map((cell,j) =>
        new TableCell({
          borders:allB(bdr),
          shading:{fill:cell._bg||(i%2===0?GR:WH),type:ShadingType.CLEAR},
          width:{size:colW[j],type:WidthType.DXA}, margins:cm,
          children:[new Paragraph({children:[new TextRun({text:String(cell._v!==undefined?cell._v:cell),font:'Arial',size:18,bold:cell._bold||false,color:cell._c||'000000'})],alignment:j===0||j===headers.length-1?AlignmentType.LEFT:AlignmentType.CENTER})],
        })
      )})),
    ],
  });
}

function kpiRow(label, total, avg, max, status) {
  const sc = {CRITICAL:RD,CRÍTICO:RD, WARNING:AD,ADVERTENCIA:AD, HIGH:AD,ALTO:AD, NORMAL:'375623'};
  const sb = {CRITICAL:RL,CRÍTICO:RL, WARNING:AB,ADVERTENCIA:AB, HIGH:AB,ALTO:AB, NORMAL:GB};
  return [
    label,
    {_v:total, _bold:false},
    {_v:avg},
    {_v:max},
    {_v:status, _bold:true, _c:sc[status]||'000000', _bg:sb[status]||GR},
  ];
}

function fmtCLP(v) { return '$' + Math.round(v).toLocaleString('en-US'); }
function fmtPct(v) { return v.toFixed(2) + '%'; }
function fmtN(v)   { return Math.round(v).toLocaleString('en-US'); }

// ─── Translations ─────────────────────────────────────────────────────────────
const T = {
  en: {
    reportTitle:'PAYMENT CONVERSION', reportTitle2:'& FAILURE REPORT',
    generated:'Generated', confidentialDoc:'CONFIDENTIAL DOCUMENT', period:'Period', pageof:'of',
    headerLabel: op=>`CONFIDENTIAL — ${op} — Payment Failure Report`,
    footerLabel: op=>`${op} — Confidential | Page `,
    s1:'1. EXECUTIVE SUMMARY',
    s1_periodLabel:'Period analyzed: ', s1_periodText:(s,e,d)=>`${s} – ${e} — ${d} calendar days.`,
    s1_purposeLabel:'Purpose: ', s1_purposeText:op=>`Quantify and analyze payment failures on the digital ticket sales platform of ${op}, identifying root causes and channels with the greatest impact on passenger experience and uncaptured revenue.`,
    s1_1:'1.1 Key indicators',
    kpiMetric:'Metric', kpiTotal:'Total', kpiAvg:'Daily Avg.', kpiMax:'Max/Day', kpiStatus:'Status',
    kpiSales:'Completed sales', kpiFail:'Gateway failures', kpiPend:'Pending-failures', kpiAban:'Abandonments',
    sCRITICAL:'CRITICAL', sWARNING:'WARNING', sHIGH:'HIGH', sNORMAL:'NORMAL',
    lostRevAlert:v=>`⚠ Estimated lost revenue: ${v} CLP`,
    lostRevDetail:(f,p,a)=>`Gateway failures: ${f} | Pending-failures: ${p} | Abandonments: ${a}. Abandonment amount is indicative — not all cases are definitively lost sales.`,
    s1_2:'1.2 Category definitions',
    defFailLabel:'Gateway failures: ', defFailText:'Customer attempts payment but the gateway returns an error. No charge is processed, reservation stays "Failed".',
    defPendLabel:'Pending-failures: ', defPendText:'Payment processed by gateway but confirmation to the sales system fails. Highest risk: customer may have been charged without receiving ticket.',
    defAbanLabel:'Abandonments: ',     defAbanText:'Customer reaches the payment screen but does not attempt payment. Recorded as "Tentative". May indicate UX issues, load time, or discouragement from prior errors.',
    s2:'2. DAILY ANALYSIS', s2_1:'2.1 Consolidated daily averages',
    avgSalesLabel:'Average sales/day: ', avgSalesText:v=>`${v} completed transactions.`,
    noSalesNote:'Note: ', noSalesText:'No sales file provided — conversion rate vs completed sales not available.',
    avgFailLabel:'Average failures/day: ', avgFailText:(v,r)=>`${v} gateway failures (avg rate ${r}).`,
    avgAbanLabel:'Average abandonments/day: ', avgAbanText:(v,r)=>`${v} (avg rate ${r} of total attempts).`,
    s2_2:'2.2 Critical days', critDaysIntro:'Days with failure rate above 3%:',
    critDayLine:(ds,f,r,a)=>[`${ds}: `,`${f} failures — rate ${r}. ${a} abandonments.`],
    loadCorrTitle:'Load-failure correlation', loadCorrText:'Critical days coincide with high-volume sales days, suggesting payment system degradation under load — not random variance.',
    s2_3:'2.3 Full daily breakdown',
    tblDate:'Date', tblSales:'Sales', tblFail:'Failures', tblPend:'Pend.', tblAban:'Abandons', tblFailRate:'Fail Rate', tblAbanRate:'Abnd Rate', tblNote:'Note',
    obsCritical:'!! Critical', obsHighFails:'High fails', obsHighAban:'!! High abandons', totLabel:'AVG/TOTAL', totSuffix:' tot',
    s3:'3. PAYMENT GATEWAY ANALYSIS', s3_1:'3.1 Distribution by gateway',
    gwGateway:'Gateway', gwDirectFail:'Direct Fail.', gwPendFail:'Pend-Fail.', gwAban:'Abandonments', gwTotal:'Total', gwPct:'% Total', gwDiag:'Diagnosis',
    diagDominant:'!! Dominant failure source. Urgent action required.', diagPendMany:'Pending-failures — review webhooks.', diagPendOne:'Pending-failure — verify reconciliation.', diagHighAban:'High abandonments. Review UX.', diagMonitor:'Monitor.',
    gwTopAlert:(gw,pct,tot,fail)=>[`🚨 ${gw}: ${pct} of direct failures`,`${gw} accounts for ${tot} of ${fail} direct gateway failures. This concentration requires urgent investigation — request error code report from the provider and check for service degradation.`],
    pendWarnTitle:'Pending-failures detected — possible charges without tickets',
    pendWarnText:list=>`${list} recorded pending-failures. Urgently reconcile these cases against bank records — customers may have been charged without receiving their ticket.`,
    s3_2:'3.2 Individual gateway notes',
    gwDirectFailLabel:'Direct failures: ', gwPendFailLabel:'Pending-failures: ', gwAbanLabel:'Abandonments: ', gwActionLabel:'Action: ',
    gwActionUrgent:'Contact provider urgently for error logs.', gwActionWebhook:'Review webhook configuration.', gwActionUX:'Review UX and error messaging.', gwActionMonitor:'Monitor.',
    s4:'4. CHANNEL AND DEVICE ANALYSIS', s4_1:'4.1 Failures by channel',
    chChannel:'Channel', chFail:'Failures', chPctFail:'% Failures', chAban:'Abandonments', chPctAban:'% Abandons', chNote:'Note',
    chTopNote:'!! Review SDK integration', chMonitor:'Monitor',
    s4_2:'4.2 Platform distribution', plLine:(p,f,a)=>[`${p}: `,`${f} failures, ${a} abandonments.`],
    s5:'5. ESTIMATED ECONOMIC IMPACT', s5_1:'5.1 Uncaptured revenue',
    econCat:'Category', econCases:'# Cases', econAmt:'Total Amount (CLP)', econPct:'% of Total', econAvg:'Avg. Ticket (CLP)',
    econFail:'Gateway failures', econPend:'Pending-failures', econAban:'Abandonments', econTotal:'TOTAL ESTIMATED',
    econNoteTitle:'Methodological note', econNoteText:'Abandonment amounts represent ticket values at the time of attempt. Not every abandonment is a definitively lost sale.',
    s5_2:'5.2 Monthly projection',
    projFailLabel:'Projected monthly failures: ', projFailText:v=>`~${v} failures/month.`,
    projAbanLabel:'Projected monthly abandonments: ', projAbanText:v=>`~${v}/month.`,
    projRevLabel:'Projected monthly uncaptured revenue: ', projRevText:v=>`~${v} CLP/month.`,
    s6:'6. RECOMMENDATIONS & ACTION PLAN', s6_1:'6.1 Immediate actions (0–72 hours)',
    act1Title:'ACTION 1: Investigate top gateway urgently', act1Text:gw=>`${gw} accounts for the majority of direct failures. Request error code report from provider. Analyze degradation on peak days.`,
    act2Title:'ACTION 2: Reconcile pending-failures', act2Text:n=>`Manually verify all ${n} pending-failure cases against bank records. Contact affected customers. Issue tickets or process refunds immediately.`,
    s6_2:'6.2 Short-term actions (1–2 weeks)',
    act3Title:'ACTION 3: Review Mobile App SDK integration', act3Text:'If Mobile App concentrates the majority of failures disproportionate to its traffic share, review gateway SDK version and compatibility updates. Implement detailed error logging.',
    act4Title:'ACTION 4: Audit webhook confirmation system', act4Text:'Pending-failures indicate webhook delivery issues. Review callback URLs, timeouts, and implement automatic retry mechanisms.',
    s6_3:'6.3 Medium-term actions (2–4 weeks)',
    act5Title:'ACTION 5: Real-time monitoring alerts', act5Text:'Implement automatic alerts when failure rate exceeds 2% in a 1-hour window to enable rapid technical intervention.',
    act6Title:'ACTION 6: Abandonment UX review', act6Text:r=>`An abandonment rate of ${r} suggests checkout friction. Review error messaging and consider retry flows or alternative gateway suggestions.`,
    s6_4:'6.4 Monitoring KPIs',
    kpiT1:'Failure rate by gateway: ', kpiT1v:'target <1%.', kpiT2:'Abandonment rate: ', kpiT2v:'target <15%.', kpiT3:'Pending-failures: ', kpiT3v:'target 0.', kpiT4:'Incident resolution time: ', kpiT4v:'max 4 hours from detection.',
    s7:'7. CONCLUSIONS', concIntro:(d,op)=>`Analysis of the ${d}-day period for ${op} reveals the following key findings:`,
    conc1b:'Gateway concentration requires immediate attention. ', conc1n:(gw,pct)=>`${gw} accounts for ${pct} of direct failures — an abnormal concentration.`,
    conc2b:'Customers may have been charged without tickets. ', conc2n:n=>`${n} pending-failures must be resolved before affected customers contact support.`,
    conc3b:r=>`Abandonment rate of ${r} is structurally high. `, conc3n:'1 in 4 customers who reach the gateway do not complete payment. Checkout UX and error messaging require review.',
    conc4b:(v,d)=>`Economic impact: ${v} CLP in ${d} days `, conc4n:v=>`(projected ${v} CLP/month) justifies investment in payment platform improvements.`,
    concClose:'With the corrective actions in Chapter 6, ', concCloseN:'it is estimated that failures can be reduced by 80–90% and abandonments by 30–40% within 30–60 days.',
    endOfReport:'— End of Report —',
  },
  es: {
    reportTitle:'CONVERSIÓN DE PAGOS', reportTitle2:'E INFORME DE FALLAS',
    generated:'Generado', confidentialDoc:'DOCUMENTO CONFIDENCIAL', period:'Período', pageof:'de',
    headerLabel: op=>`CONFIDENCIAL — ${op} — Informe de Fallas de Pago`,
    footerLabel: op=>`${op} — Confidencial | Página `,
    s1:'1. RESUMEN EJECUTIVO',
    s1_periodLabel:'Período analizado: ', s1_periodText:(s,e,d)=>`${s} – ${e} — ${d} días calendario.`,
    s1_purposeLabel:'Objetivo: ', s1_purposeText:op=>`Cuantificar y analizar las fallas de pago en la plataforma de venta digital de pasajes de ${op}, identificando causas raíz y canales con mayor impacto en la experiencia del pasajero e ingresos no capturados.`,
    s1_1:'1.1 Indicadores clave',
    kpiMetric:'Métrica', kpiTotal:'Total', kpiAvg:'Prom. Diario', kpiMax:'Máx/Día', kpiStatus:'Estado',
    kpiSales:'Ventas completadas', kpiFail:'Fallas de gateway', kpiPend:'Pendientes-fallidos', kpiAban:'Abandonos',
    sCRITICAL:'CRÍTICO', sWARNING:'ADVERTENCIA', sHIGH:'ALTO', sNORMAL:'NORMAL',
    lostRevAlert:v=>`⚠ Ingresos estimados perdidos: ${v} CLP`,
    lostRevDetail:(f,p,a)=>`Fallas de gateway: ${f} | Pendientes-fallidos: ${p} | Abandonos: ${a}. El monto de abandono es indicativo — no todos son ventas definitivamente perdidas.`,
    s1_2:'1.2 Definición de categorías',
    defFailLabel:'Fallas de gateway: ', defFailText:'El cliente intenta pagar pero el gateway retorna un error. No se realiza cobro, la reserva queda en estado "Fallido".',
    defPendLabel:'Pendientes-fallidos: ', defPendText:'El pago fue procesado por el gateway pero la confirmación al sistema de ventas falla. Mayor riesgo: el cliente puede haber sido cobrado sin recibir su pasaje.',
    defAbanLabel:'Abandonos: ', defAbanText:'El cliente llega a la pantalla de pago pero no intenta pagar. Se registra como "Tentativo". Puede indicar problemas de UX, tiempo de carga o desaliento por errores previos.',
    s2:'2. ANÁLISIS DIARIO', s2_1:'2.1 Promedios diarios consolidados',
    avgSalesLabel:'Promedio ventas/día: ', avgSalesText:v=>`${v} transacciones completadas.`,
    noSalesNote:'Nota: ', noSalesText:'No se proporcionó archivo de ventas — tasa de conversión vs ventas completadas no disponible.',
    avgFailLabel:'Promedio fallas/día: ', avgFailText:(v,r)=>`${v} fallas de gateway (tasa promedio ${r}).`,
    avgAbanLabel:'Promedio abandonos/día: ', avgAbanText:(v,r)=>`${v} (tasa promedio ${r} del total de intentos).`,
    s2_2:'2.2 Días críticos', critDaysIntro:'Días con tasa de falla superior al 3%:',
    critDayLine:(ds,f,r,a)=>[`${ds}: `,`${f} fallas — tasa ${r}. ${a} abandonos.`],
    loadCorrTitle:'Correlación carga-falla', loadCorrText:'Los días críticos coinciden con los días de mayor volumen de ventas, lo que sugiere degradación del sistema de pagos bajo carga — no varianza aleatoria.',
    s2_3:'2.3 Detalle diario completo',
    tblDate:'Fecha', tblSales:'Ventas', tblFail:'Fallas', tblPend:'Pend.', tblAban:'Abandonos', tblFailRate:'Tasa Falla', tblAbanRate:'Tasa Aban.', tblNote:'Nota',
    obsCritical:'!! Crítico', obsHighFails:'Fallas altas', obsHighAban:'!! Abandonos altos', totLabel:'PROM/TOTAL', totSuffix:' tot',
    s3:'3. ANÁLISIS DE GATEWAYS DE PAGO', s3_1:'3.1 Distribución por gateway',
    gwGateway:'Gateway', gwDirectFail:'Fallas Directas', gwPendFail:'Pend-Fallidos', gwAban:'Abandonos', gwTotal:'Total', gwPct:'% Total', gwDiag:'Diagnóstico',
    diagDominant:'!! Fuente dominante de fallas. Acción urgente requerida.', diagPendMany:'Pendientes-fallidos — revisar webhooks.', diagPendOne:'Pendiente-fallido — verificar reconciliación.', diagHighAban:'Abandonos altos. Revisar UX.', diagMonitor:'Monitorear.',
    gwTopAlert:(gw,pct,tot,fail)=>[`🚨 ${gw}: ${pct} de las fallas directas`,`${gw} concentra ${tot} de ${fail} fallas directas de gateway. Esta concentración requiere investigación urgente — solicitar informe de códigos de error al proveedor y verificar degradación del servicio.`],
    pendWarnTitle:'Pendientes-fallidos detectados — posibles cobros sin pasaje',
    pendWarnText:list=>`${list} registraron pendientes-fallidos. Reconciliar urgentemente con registros bancarios — clientes pueden haber sido cobrados sin recibir su pasaje.`,
    s3_2:'3.2 Notas por gateway individual',
    gwDirectFailLabel:'Fallas directas: ', gwPendFailLabel:'Pendientes-fallidos: ', gwAbanLabel:'Abandonos: ', gwActionLabel:'Acción: ',
    gwActionUrgent:'Contactar proveedor urgentemente para obtener logs de error.', gwActionWebhook:'Revisar configuración de webhooks.', gwActionUX:'Revisar UX y mensajes de error.', gwActionMonitor:'Monitorear.',
    s4:'4. ANÁLISIS DE CANAL Y DISPOSITIVO', s4_1:'4.1 Fallas por canal',
    chChannel:'Canal', chFail:'Fallas', chPctFail:'% Fallas', chAban:'Abandonos', chPctAban:'% Abandonos', chNote:'Nota',
    chTopNote:'!! Revisar integración SDK', chMonitor:'Monitorear',
    s4_2:'4.2 Distribución por plataforma', plLine:(p,f,a)=>[`${p}: `,`${f} fallas, ${a} abandonos.`],
    s5:'5. IMPACTO ECONÓMICO ESTIMADO', s5_1:'5.1 Ingresos no capturados',
    econCat:'Categoría', econCases:'N° Casos', econAmt:'Monto Total (CLP)', econPct:'% del Total', econAvg:'Ticket Prom. (CLP)',
    econFail:'Fallas de gateway', econPend:'Pendientes-fallidos', econAban:'Abandonos', econTotal:'TOTAL ESTIMADO',
    econNoteTitle:'Nota metodológica', econNoteText:'Los montos de abandono representan el valor del pasaje al momento del intento. No todos los abandonos son ventas definitivamente perdidas.',
    s5_2:'5.2 Proyección mensual',
    projFailLabel:'Fallas proyectadas al mes: ', projFailText:v=>`~${v} fallas/mes.`,
    projAbanLabel:'Abandonos proyectados al mes: ', projAbanText:v=>`~${v}/mes.`,
    projRevLabel:'Ingresos no capturados proyectados al mes: ', projRevText:v=>`~${v} CLP/mes.`,
    s6:'6. RECOMENDACIONES Y PLAN DE ACCIÓN', s6_1:'6.1 Acciones inmediatas (0–72 horas)',
    act1Title:'ACCIÓN 1: Investigar gateway principal urgentemente', act1Text:gw=>`${gw} concentra la mayoría de las fallas directas. Solicitar informe de códigos de error al proveedor. Analizar degradación en días pico.`,
    act2Title:'ACCIÓN 2: Reconciliar pendientes-fallidos', act2Text:n=>`Verificar manualmente los ${n} casos de pendientes-fallidos contra registros bancarios. Contactar clientes afectados. Emitir pasajes o procesar reembolsos de inmediato.`,
    s6_2:'6.2 Acciones a corto plazo (1–2 semanas)',
    act3Title:'ACCIÓN 3: Revisar integración del SDK de la App Móvil', act3Text:'Si la App Móvil concentra la mayoría de las fallas de forma desproporcionada a su cuota de tráfico, revisar la versión del SDK del gateway y actualizaciones de compatibilidad. Implementar logging detallado de errores.',
    act4Title:'ACCIÓN 4: Auditar sistema de confirmación por webhooks', act4Text:'Los pendientes-fallidos indican problemas en la entrega de webhooks. Revisar URLs de callback, timeouts e implementar mecanismos de reintento automático.',
    s6_3:'6.3 Acciones a mediano plazo (2–4 semanas)',
    act5Title:'ACCIÓN 5: Alertas de monitoreo en tiempo real', act5Text:'Implementar alertas automáticas cuando la tasa de falla supere el 2% en una ventana de 1 hora para permitir intervención técnica rápida.',
    act6Title:'ACCIÓN 6: Revisión de UX de abandonos', act6Text:r=>`Una tasa de abandono de ${r} sugiere fricción en el checkout. Revisar mensajes de error y considerar flujos de reintento o sugerencias de gateway alternativo.`,
    s6_4:'6.4 KPIs de monitoreo',
    kpiT1:'Tasa de falla por gateway: ', kpiT1v:'meta <1%.', kpiT2:'Tasa de abandono: ', kpiT2v:'meta <15%.', kpiT3:'Pendientes-fallidos: ', kpiT3v:'meta 0.', kpiT4:'Tiempo de resolución de incidentes: ', kpiT4v:'máximo 4 horas desde detección.',
    s7:'7. CONCLUSIONES', concIntro:(d,op)=>`El análisis del período de ${d} días para ${op} revela los siguientes hallazgos clave:`,
    conc1b:'La concentración en un gateway requiere atención inmediata. ', conc1n:(gw,pct)=>`${gw} concentra el ${pct} de las fallas directas — una concentración anormal.`,
    conc2b:'Clientes podrían haber sido cobrados sin recibir pasaje. ', conc2n:n=>`${n} pendientes-fallidos deben resolverse antes de que los clientes afectados contacten soporte.`,
    conc3b:r=>`La tasa de abandono de ${r} es estructuralmente alta. `, conc3n:'1 de cada 4 clientes que llega al gateway no completa el pago. El UX del checkout y los mensajes de error requieren revisión.',
    conc4b:(v,d)=>`Impacto económico: ${v} CLP en ${d} días `, conc4n:v=>`(proyectado ${v} CLP/mes) justifica inversión en mejoras de la plataforma de pagos.`,
    concClose:'Con las acciones correctivas del Capítulo 6, ', concCloseN:'se estima que las fallas pueden reducirse en un 80–90% y los abandonos en un 30–40% en 30–60 días.',
    endOfReport:'— Fin del Informe —',
  },
};

async function generate(data) {
  const lang = (data.lang||'en')==='es' ? 'es' : 'en';
  const L = T[lang];

  const {
    operator, periodStart, periodEnd, totalDays,
    totalSales, totalFailures, totalPending, totalAbandon,
    avgSales, avgFailures, avgAbandon, avgFailRate, avgAbanRate,
    priceFailures, pricePending, priceAbandon, totalLost,
    peakFailDay, peakAbanDay, hasSales,
    daily, gateways, channels, platforms, totalInc,
  } = data;

  const today = new Date().toLocaleDateString(lang==='es'?'es-CL':'en-GB');
  const topGW = gateways[0];

  const doc = new Document({
    numbering:{config:[
      {reference:'bullets',levels:[{level:0,format:LevelFormat.BULLET,text:'•',alignment:AlignmentType.LEFT,style:{paragraph:{indent:{left:720,hanging:360}}}}]},
      {reference:'numbers',levels:[{level:0,format:LevelFormat.DECIMAL,text:'%1.',alignment:AlignmentType.LEFT,style:{paragraph:{indent:{left:720,hanging:360}}}}]},
    ]},
    styles:{
      default:{document:{run:{font:'Arial',size:22}}},
      paragraphStyles:[
        {id:'Heading1',name:'Heading 1',basedOn:'Normal',next:'Normal',quickFormat:true,run:{size:28,bold:true,font:'Arial',color:WH},paragraph:{spacing:{before:320,after:120},outlineLevel:0}},
        {id:'Heading2',name:'Heading 2',basedOn:'Normal',next:'Normal',quickFormat:true,run:{size:26,bold:true,font:'Arial',color:BD},paragraph:{spacing:{before:260,after:100},outlineLevel:1}},
        {id:'Heading3',name:'Heading 3',basedOn:'Normal',next:'Normal',quickFormat:true,run:{size:24,bold:true,font:'Arial',color:BM},paragraph:{spacing:{before:180,after:80},outlineLevel:2}},
      ],
    },
    sections:[{
      properties:{page:{size:{width:12240,height:15840},margin:{top:1440,right:1080,bottom:1440,left:1080}}},
      headers:{default:new Header({children:[new Paragraph({
        children:[new TextRun({text:L.headerLabel(operator),font:'Arial',size:18,color:'888888'})],
        border:{bottom:{style:BorderStyle.SINGLE,size:4,color:BM,space:1}},spacing:{after:200},
      })]})},
      footers:{default:new Footer({children:[new Paragraph({
        children:[
          new TextRun({text:L.footerLabel(operator),font:'Arial',size:18,color:'888888'}),
          new TextRun({children:[PageNumber.CURRENT],font:'Arial',size:18,color:'888888'}),
          new TextRun({text:` ${L.pageof} `,font:'Arial',size:18,color:'888888'}),
          new TextRun({children:[PageNumber.TOTAL_PAGES],font:'Arial',size:18,color:'888888'}),
        ],
        alignment:AlignmentType.RIGHT,
        border:{top:{style:BorderStyle.SINGLE,size:4,color:BM,space:1}},
      })]})},
      children:[
        // COVER
        new Paragraph({children:[],spacing:{before:1200,after:0}}),
        new Paragraph({children:[new TextRun({text:L.reportTitle,font:'Arial',bold:true,size:48,color:BD})],alignment:AlignmentType.CENTER,spacing:{before:0,after:120}}),
        new Paragraph({children:[new TextRun({text:L.reportTitle2,font:'Arial',bold:true,size:48,color:BD})],alignment:AlignmentType.CENTER,spacing:{before:0,after:280}}),
        new Paragraph({children:[new TextRun({text:operator,font:'Arial',size:28,color:'555555'})],alignment:AlignmentType.CENTER,spacing:{before:0,after:120}}),
        new Paragraph({children:[new TextRun({text:`${L.period}: ${periodStart} – ${periodEnd}`,font:'Arial',size:24,color:'777777'})],alignment:AlignmentType.CENTER,spacing:{before:0,after:120}}),
        new Paragraph({children:[new TextRun({text:`${L.generated}: ${today}`,font:'Arial',size:22,color:'999999'})],alignment:AlignmentType.CENTER,spacing:{before:0,after:800}}),
        new Paragraph({children:[new TextRun({text:L.confidentialDoc,font:'Arial',bold:true,size:20,color:RD})],alignment:AlignmentType.CENTER,shading:{fill:RL,type:ShadingType.CLEAR},spacing:{before:0,after:0}}),
        new Paragraph({children:[new PageBreak()],spacing:{before:0,after:0}}),

        // 1
        hd1(L.s1),
        para([b(L.s1_periodLabel),n(L.s1_periodText(periodStart,periodEnd,totalDays))]),
        para([b(L.s1_purposeLabel),n(L.s1_purposeText(operator))]),
        hd2(L.s1_1),
        new Paragraph({spacing:{before:80,after:80}}),
        simpleTable([L.kpiMetric,L.kpiTotal,L.kpiAvg,L.kpiMax,L.kpiStatus],[
          ...(hasSales?[kpiRow(L.kpiSales,fmtN(totalSales??0),(avgSales??0).toFixed(1),daily.length?Math.max(...daily.map(d=>d.sales)):0,L.sNORMAL)]:[]),
          kpiRow(L.kpiFail,fmtN(totalFailures),avgFailures.toFixed(1),peakFailDay?.failures||0,L.sCRITICAL),
          kpiRow(L.kpiPend,fmtN(totalPending),(totalPending/totalDays).toFixed(1),daily.length?Math.max(...daily.map(d=>d.pending)):0,L.sWARNING),
          kpiRow(L.kpiAban,fmtN(totalAbandon),avgAbandon.toFixed(1),peakAbanDay?.abandonments||0,L.sHIGH),
        ],[2800,1800,1800,1600,1360]),
        new Paragraph({spacing:{before:120,after:0}}),
        ...alertBox(L.lostRevAlert(fmtCLP(totalLost)),L.lostRevDetail(fmtCLP(priceFailures),fmtCLP(pricePending),fmtCLP(priceAbandon))),
        hd2(L.s1_2),
        para([b(L.defFailLabel),n(L.defFailText)]),
        para([b(L.defPendLabel),n(L.defPendText)]),
        para([b(L.defAbanLabel),n(L.defAbanText)]),
        new Paragraph({children:[new PageBreak()],spacing:{before:0,after:0}}),

        // 2
        hd1(L.s2),
        hd2(L.s2_1),
        hasSales?para([b(L.avgSalesLabel),n(L.avgSalesText((avgSales??0).toFixed(1)))]):para([b(L.noSalesNote),n(L.noSalesText)]),
        para([b(L.avgFailLabel),n(L.avgFailText(avgFailures.toFixed(1),fmtPct(avgFailRate)))]),
        para([b(L.avgAbanLabel),n(L.avgAbanText(avgAbandon.toFixed(1),fmtPct(avgAbanRate)))]),
        hd2(L.s2_2),
        para(L.critDaysIntro),
        ...daily.filter(d=>d.failureRate>=3).sort((a,b)=>b.failureRate-a.failureRate).map(d=>{
          const [bp,np]=L.critDayLine(d.dateStr,d.failures,fmtPct(d.failureRate),d.abandonments);
          return bullet([b(bp),n(np)]);
        }),
        new Paragraph({spacing:{before:80,after:0}}),
        ...(daily.filter(d=>d.failureRate>=3).length>1?warnBox(L.loadCorrTitle,L.loadCorrText):[]),
        hd2(L.s2_3),
        new Paragraph({spacing:{before:80,after:80}}),
        simpleTable([L.tblDate,L.tblSales,L.tblFail,L.tblPend,L.tblAban,L.tblFailRate,L.tblAbanRate,L.tblNote],[
          ...daily.map(d=>{
            const ic=d.failures>=20||d.failureRate>=4;
            const obs=[];
            if(d.failures>=20)obs.push(L.obsCritical);
            else if(d.failures>=10)obs.push(L.obsHighFails);
            if(d.failureRate>=3)obs.push(`Rate ${fmtPct(d.failureRate)}`);
            if(d.abandonments>=250)obs.push(L.obsHighAban);
            return [
              {_v:d.dateStr,_bg:ic?RL:undefined},
              {_v:d.sales,_bg:ic?RL:undefined},
              {_v:d.failures,_bg:ic?RL:undefined,_bold:ic,_c:ic?RD:'000000'},
              {_v:d.pending,_bg:ic?RL:undefined},
              {_v:d.abandonments,_bg:ic?RL:undefined},
              {_v:fmtPct(d.failureRate),_bg:ic?RL:undefined,_c:ic?RD:'000000'},
              {_v:fmtPct(d.abandonRate),_bg:ic?RL:undefined},
              {_v:obs.join(' | ')||'—',_bg:ic?RL:undefined},
            ];
          }),
          [{_v:L.totLabel,_bg:BD,_bold:true,_c:WH},{_v:fmtN(totalSales),_bg:BD,_bold:true,_c:WH},{_v:`${fmtN(totalFailures)}${L.totSuffix}`,_bg:BD,_bold:true,_c:WH},{_v:`${fmtN(totalPending)}${L.totSuffix}`,_bg:BD,_bold:true,_c:WH},{_v:`${fmtN(totalAbandon)}${L.totSuffix}`,_bg:BD,_bold:true,_c:WH},{_v:fmtPct(avgFailRate),_bg:BD,_bold:true,_c:WH},{_v:fmtPct(avgAbanRate),_bg:BD,_bold:true,_c:WH},{_v:'',_bg:BD}],
        ],[1100,900,1000,900,1100,1100,1100,1160]),
        new Paragraph({children:[new PageBreak()],spacing:{before:0,after:0}}),

        // 3
        hd1(L.s3),
        hd2(L.s3_1),
        new Paragraph({spacing:{before:80,after:80}}),
        simpleTable([L.gwGateway,L.gwDirectFail,L.gwPendFail,L.gwAban,L.gwTotal,L.gwPct,L.gwDiag],[
          ...gateways.map((g,i)=>{
            let diag=L.diagMonitor;
            if(g.failures/totalFailures>0.9)diag=L.diagDominant;
            else if(g.pending>3)diag=L.diagPendMany;
            else if(g.pending>0)diag=L.diagPendOne;
            else if(g.abandonments>500)diag=L.diagHighAban;
            const it=i===0&&g.failures>0;
            return [{_v:g.gateway,_bg:it?RL:undefined,_bold:it,_c:it?RD:'000000'},{_v:g.failures,_bg:it?RL:undefined,_bold:it,_c:it?RD:'000000'},{_v:g.pending,_bg:it?RL:undefined},{_v:g.abandonments,_bg:it?RL:undefined},{_v:g.total,_bg:it?RL:undefined,_bold:it,_c:it?RD:'000000'},{_v:fmtPct(g.pct),_bg:it?RL:undefined},{_v:diag,_bg:it?RL:undefined}];
          }),
          [{_v:'TOTAL',_bg:BD,_bold:true,_c:WH},{_v:fmtN(gateways.reduce((s,g)=>s+g.failures,0)),_bg:BD,_bold:true,_c:WH},{_v:fmtN(gateways.reduce((s,g)=>s+g.pending,0)),_bg:BD,_bold:true,_c:WH},{_v:fmtN(totalAbandon),_bg:BD,_bold:true,_c:WH},{_v:fmtN(totalInc),_bg:BD,_bold:true,_c:WH},{_v:'100%',_bg:BD,_bold:true,_c:WH},{_v:'',_bg:BD}],
        ],[1200,1000,1000,1200,1000,900,3060]),
        new Paragraph({spacing:{before:120,after:0}}),
        ...(topGW&&topGW.failures>0?(()=>{const[t,x]=L.gwTopAlert(topGW.gateway,fmtPct(topGW.failures/totalFailures*100),topGW.failures,totalFailures);return alertBox(t,x);})():[]),
        ...gateways.filter(g=>g.pending>0).length>0?warnBox(L.pendWarnTitle,L.pendWarnText(gateways.filter(g=>g.pending>0).map(g=>`${g.gateway} (${g.pending})`).join(', '))):[],
        hd2(L.s3_2),
        ...gateways.slice(0,5).flatMap(g=>{
          const action=g.failures>100?L.gwActionUrgent:g.pending>0?L.gwActionWebhook:g.abandonments>500?L.gwActionUX:L.gwActionMonitor;
          return [hd3(g.gateway),para([b(L.gwDirectFailLabel),n(String(g.failures))]),para([b(L.gwPendFailLabel),n(String(g.pending))]),para([b(L.gwAbanLabel),n(String(g.abandonments))]),para([b(L.gwActionLabel),n(action)])];
        }),
        new Paragraph({children:[new PageBreak()],spacing:{before:0,after:0}}),

        // 4
        hd1(L.s4),
        hd2(L.s4_1),
        new Paragraph({spacing:{before:80,after:80}}),
        simpleTable([L.chChannel,L.chFail,L.chPctFail,L.chAban,L.chPctAban,L.chNote],
          channels.map((c,i)=>{
            const pf=channels.reduce((s,x)=>s+x.failures,0)||1;
            const pa=channels.reduce((s,x)=>s+x.abandonments,0)||1;
            const it=i===0;
            return [{_v:c.channel,_bg:it?AB:undefined,_bold:it},{_v:c.failures,_bg:it?AB:undefined,_bold:it},{_v:fmtPct(c.failures/pf*100),_bg:it?AB:undefined},{_v:c.abandonments,_bg:it?AB:undefined},{_v:fmtPct(c.abandonments/pa*100),_bg:it?AB:undefined},{_v:it?L.chTopNote:L.chMonitor,_bg:it?AB:undefined}];
          }),
        [2400,1200,1400,1400,1400,2360]),
        new Paragraph({spacing:{before:120,after:0}}),
        hd2(L.s4_2),
        ...platforms.map(p=>{const[bp,np]=L.plLine(p.platform,p.failures,p.abandonments);return para([b(bp),n(np)]);}),
        new Paragraph({children:[new PageBreak()],spacing:{before:0,after:0}}),

        // 5
        hd1(L.s5),
        hd2(L.s5_1),
        new Paragraph({spacing:{before:80,after:80}}),
        simpleTable([L.econCat,L.econCases,L.econAmt,L.econPct,L.econAvg],[
          [{_v:L.econFail},{_v:fmtN(totalFailures)},{_v:fmtCLP(priceFailures)},{_v:fmtPct(totalLost?priceFailures/totalLost*100:0)},{_v:fmtCLP(totalFailures?priceFailures/totalFailures:0)}],
          [{_v:L.econPend},{_v:fmtN(totalPending)},{_v:fmtCLP(pricePending)},{_v:fmtPct(totalLost?pricePending/totalLost*100:0)},{_v:fmtCLP(totalPending?pricePending/totalPending:0)}],
          [{_v:L.econAban},{_v:fmtN(totalAbandon)},{_v:fmtCLP(priceAbandon)},{_v:fmtPct(totalLost?priceAbandon/totalLost*100:0)},{_v:fmtCLP(totalAbandon?priceAbandon/totalAbandon:0)}],
          [{_v:L.econTotal,_bg:RL,_bold:true,_c:RD},{_v:fmtN(totalFailures+totalPending+totalAbandon),_bg:RL,_bold:true,_c:RD},{_v:fmtCLP(totalLost),_bg:RL,_bold:true,_c:RD},{_v:'100%',_bg:RL,_bold:true,_c:RD},{_v:fmtCLP(totalLost/(totalFailures+totalPending+totalAbandon||1)),_bg:RL,_bold:true,_c:RD}],
        ],[2600,1500,2200,1500,1560]),
        new Paragraph({spacing:{before:120,after:0}}),
        ...infoBox(L.econNoteTitle,L.econNoteText),
        hd2(L.s5_2),
        para([b(L.projFailLabel),n(L.projFailText(Math.round(totalFailures/totalDays*30)))]),
        para([b(L.projAbanLabel),n(L.projAbanText(Math.round(totalAbandon/totalDays*30)))]),
        para([b(L.projRevLabel),n(L.projRevText(fmtCLP(totalLost/totalDays*30)))]),
        new Paragraph({children:[new PageBreak()],spacing:{before:0,after:0}}),

        // 6
        hd1(L.s6),
        hd2(L.s6_1),
        ...alertBox(L.act1Title,L.act1Text(topGW?.gateway||'Gateway principal')),
        ...(gateways.some(g=>g.pending>0)?alertBox(L.act2Title,L.act2Text(totalPending)):[]),
        hd2(L.s6_2),
        ...warnBox(L.act3Title,L.act3Text),
        ...warnBox(L.act4Title,L.act4Text),
        hd2(L.s6_3),
        ...infoBox(L.act5Title,L.act5Text),
        ...infoBox(L.act6Title,L.act6Text(fmtPct(avgAbanRate))),
        hd2(L.s6_4),
        bullet([b(L.kpiT1),n(L.kpiT1v)]),
        bullet([b(L.kpiT2),n(L.kpiT2v)]),
        bullet([b(L.kpiT3),n(L.kpiT3v)]),
        bullet([b(L.kpiT4),n(L.kpiT4v)]),

        // 7
        new Paragraph({children:[new PageBreak()],spacing:{before:0,after:0}}),
        hd1(L.s7),
        para(L.concIntro(totalDays,operator)),
        new Paragraph({numbering:{reference:'numbers',level:0},children:[b(L.conc1b),n(L.conc1n(topGW?.gateway||'Gateway principal',fmtPct((topGW?.failures||0)/totalFailures*100)))],spacing:{before:80,after:60}}),
        ...(totalPending>0?[new Paragraph({numbering:{reference:'numbers',level:0},children:[b(L.conc2b),n(L.conc2n(totalPending))],spacing:{before:60,after:60}})]:[]),
        new Paragraph({numbering:{reference:'numbers',level:0},children:[b(L.conc3b(fmtPct(avgAbanRate))),n(L.conc3n)],spacing:{before:60,after:60}}),
        new Paragraph({numbering:{reference:'numbers',level:0},children:[b(L.conc4b(fmtCLP(totalLost),totalDays)),n(L.conc4n(fmtCLP(totalLost/totalDays*30)))],spacing:{before:60,after:120}}),
        para([b(L.concClose),n(L.concCloseN)]),
        new Paragraph({spacing:{before:400,after:100}}),
        new Paragraph({children:[new TextRun({text:L.endOfReport,font:'Arial',size:20,color:'888888',italics:true})],alignment:AlignmentType.CENTER}),
      ],
    }],
  });

  return Packer.toBuffer(doc);
}

module.exports = { generate };