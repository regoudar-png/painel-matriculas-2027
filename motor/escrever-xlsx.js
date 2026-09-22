// Escritor mínimo de .xlsx: zip próprio + OOXML com inline strings.
const zlib=require('zlib');

/* ---------------- zip ---------------- */
const TAB=(()=>{const t=new Int32Array(256);
  for(let n=0;n<256;n++){let c=n;for(let k=0;k<8;k++)c=c&1?0xEDB88320^(c>>>1):c>>>1;t[n]=c;}
  return t;})();
function crc32(buf){let c=~0;for(let i=0;i<buf.length;i++)c=TAB[(c^buf[i])&0xFF]^(c>>>8);return ~c>>>0;}
function zip(arquivos){
  const locais=[],central=[];let off=0;
  for(const f of arquivos){
    const nome=Buffer.from(f.nome,'utf8');
    const cru=Buffer.isBuffer(f.dados)?f.dados:Buffer.from(f.dados,'utf8');
    const comp=zlib.deflateRawSync(cru,{level:9});
    const crc=crc32(cru);
    const lh=Buffer.alloc(30);
    lh.writeUInt32LE(0x04034b50,0); lh.writeUInt16LE(20,4); lh.writeUInt16LE(0x0800,6);
    lh.writeUInt16LE(8,8); lh.writeUInt16LE(0,10); lh.writeUInt16LE(0x21,12);
    lh.writeUInt32LE(crc,14); lh.writeUInt32LE(comp.length,18); lh.writeUInt32LE(cru.length,22);
    lh.writeUInt16LE(nome.length,26); lh.writeUInt16LE(0,28);
    locais.push(lh,nome,comp);
    const ch=Buffer.alloc(46);
    ch.writeUInt32LE(0x02014b50,0); ch.writeUInt16LE(20,4); ch.writeUInt16LE(20,6);
    ch.writeUInt16LE(0x0800,8); ch.writeUInt16LE(8,10); ch.writeUInt16LE(0,12);
    ch.writeUInt16LE(0x21,14); ch.writeUInt32LE(crc,16);
    ch.writeUInt32LE(comp.length,20); ch.writeUInt32LE(cru.length,24);
    ch.writeUInt16LE(nome.length,28); ch.writeUInt32LE(0,38); ch.writeUInt32LE(off,42);
    central.push(ch,nome);
    off+=30+nome.length+comp.length;
  }
  const cd=Buffer.concat(central), lc=Buffer.concat(locais);
  const eo=Buffer.alloc(22);
  eo.writeUInt32LE(0x06054b50,0);
  eo.writeUInt16LE(arquivos.length,8); eo.writeUInt16LE(arquivos.length,10);
  eo.writeUInt32LE(cd.length,12); eo.writeUInt32LE(lc.length,16);
  return Buffer.concat([lc,cd,eo]);
}

/* ---------------- helpers ---------------- */
const esc=s=>String(s==null?'':s)
  .replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;')
  .replace(/"/g,'&quot;').replace(/[\x00-\x08\x0B\x0C\x0E-\x1F]/g,'');
function colName(n){let s='';n++;while(n>0){const r=(n-1)%26;s=String.fromCharCode(65+r)+s;n=(n-r-1)/26;}return s;}
const serial=iso=>{
  if(!iso)return null;
  const p=String(iso).split('-'); if(p.length!==3)return null;
  return Math.round((Date.UTC(+p[0],+p[1]-1,+p[2])-Date.UTC(1899,11,30))/86400000);
};

/* ---------------- estilos ---------------- */
// 0 padrão · 1 cabeçalho · 2 texto · 3 data · 4 moeda · 5 percentual
// 6 número · 7 ação (quebra) · 8 marcação · 9 título · 10 legenda · 11 destaque
const STYLES=`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
<numFmts count="2"><numFmt numFmtId="164" formatCode="dd/mm/yyyy"/>
<numFmt numFmtId="165" formatCode="&quot;R$&quot;\\ #,##0.00"/></numFmts>
<fonts count="6">
<font><sz val="10"/><name val="Calibri"/></font>
<font><b/><sz val="10"/><color rgb="FFFFFFFF"/><name val="Calibri"/></font>
<font><sz val="9"/><color rgb="FF6B7A88"/><name val="Calibri"/></font>
<font><b/><sz val="14"/><color rgb="FF10314F"/><name val="Calibri"/></font>
<font><b/><sz val="10"/><name val="Calibri"/></font>
<font><b/><sz val="10"/><color rgb="FF0F7A5F"/><name val="Calibri"/></font>
</fonts>
<fills count="6">
<fill><patternFill patternType="none"/></fill>
<fill><patternFill patternType="gray125"/></fill>
<fill><patternFill patternType="solid"><fgColor rgb="FF1E5CA8"/><bgColor indexed="64"/></patternFill></fill>
<fill><patternFill patternType="solid"><fgColor rgb="FFFFF4D6"/><bgColor indexed="64"/></patternFill></fill>
<fill><patternFill patternType="solid"><fgColor rgb="FFEAF6F1"/><bgColor indexed="64"/></patternFill></fill>
<fill><patternFill patternType="solid"><fgColor rgb="FFFFF1EA"/><bgColor indexed="64"/></patternFill></fill>
</fills>
<borders count="2">
<border><left/><right/><top/><bottom/><diagonal/></border>
<border><left style="thin"><color rgb="FFD5DDE5"/></left><right style="thin"><color rgb="FFD5DDE5"/></right>
<top style="thin"><color rgb="FFD5DDE5"/></top><bottom style="thin"><color rgb="FFD5DDE5"/></bottom><diagonal/></border>
</borders>
<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>
<cellXfs count="16">
<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>
<xf numFmtId="0" fontId="1" fillId="2" borderId="1" xfId="0" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1"><alignment horizontal="left" vertical="center" wrapText="1"/></xf>
<xf numFmtId="0" fontId="0" fillId="0" borderId="1" xfId="0" applyBorder="1" applyAlignment="1"><alignment vertical="top" wrapText="1"/></xf>
<xf numFmtId="164" fontId="0" fillId="0" borderId="1" xfId="0" applyNumberFormat="1" applyBorder="1" applyAlignment="1"><alignment horizontal="center" vertical="top"/></xf>
<xf numFmtId="165" fontId="0" fillId="0" borderId="1" xfId="0" applyNumberFormat="1" applyBorder="1" applyAlignment="1"><alignment horizontal="right" vertical="top"/></xf>
<xf numFmtId="10" fontId="0" fillId="0" borderId="1" xfId="0" applyNumberFormat="1" applyBorder="1" applyAlignment="1"><alignment horizontal="center" vertical="top"/></xf>
<xf numFmtId="0" fontId="0" fillId="0" borderId="1" xfId="0" applyBorder="1" applyAlignment="1"><alignment horizontal="center" vertical="top"/></xf>
<xf numFmtId="0" fontId="0" fillId="4" borderId="1" xfId="0" applyFill="1" applyBorder="1" applyAlignment="1"><alignment vertical="top" wrapText="1"/></xf>
<xf numFmtId="0" fontId="4" fillId="3" borderId="1" xfId="0" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1"><alignment horizontal="center" vertical="center"/></xf>
<xf numFmtId="0" fontId="3" fillId="0" borderId="0" xfId="0" applyFont="1"/>
<xf numFmtId="0" fontId="2" fillId="0" borderId="0" xfId="0" applyFont="1" applyAlignment="1"><alignment vertical="top" wrapText="1"/></xf>
<xf numFmtId="0" fontId="5" fillId="0" borderId="0" xfId="0" applyFont="1"/>
<xf numFmtId="0" fontId="0" fillId="5" borderId="1" xfId="0" applyFill="1" applyBorder="1" applyAlignment="1"><alignment vertical="top" wrapText="1"/></xf>
<xf numFmtId="164" fontId="0" fillId="5" borderId="1" xfId="0" applyNumberFormat="1" applyFill="1" applyBorder="1" applyAlignment="1"><alignment horizontal="center" vertical="top"/></xf>
<xf numFmtId="165" fontId="0" fillId="5" borderId="1" xfId="0" applyNumberFormat="1" applyFill="1" applyBorder="1" applyAlignment="1"><alignment horizontal="right" vertical="top"/></xf>
<xf numFmtId="0" fontId="0" fillId="5" borderId="1" xfId="0" applyFill="1" applyBorder="1" applyAlignment="1"><alignment horizontal="center" vertical="top"/></xf>
</cellXfs>
<dxfs count="1"><dxf><font><color rgb="FF0F7A5F"/><b/></font>
<fill><patternFill><bgColor rgb="FFDDF2E9"/></patternFill></fill></dxf></dxfs>
</styleSheet>`;

/* ---------------- planilha ---------------- */
// linhas: [{cells:[{v,t,s}], h?:altura}]  t: 's'|'n'|'d'|'m'|'p'
function folha(cfg){
  const {cols=[],linhas=[],congelar=0,filtro=null,validacao=null,formatacao=null}=cfg;
  let sd='';
  linhas.forEach((ln,ri)=>{
    const r=ri+1;
    let cs='';
    (ln.cells||[]).forEach((c,ci)=>{
      if(c==null)return;
      const ref=colName(ci)+r;
      const s=c.s||0;
      if(c.v==null||c.v===''){ cs+='<c r="'+ref+'" s="'+s+'"/>'; return; }
      if(c.t==='n'||c.t==='m'||c.t==='p'){ cs+='<c r="'+ref+'" s="'+s+'"><v>'+c.v+'</v></c>'; }
      else if(c.t==='d'){ const sv=serial(c.v); cs+= sv==null?'<c r="'+ref+'" s="'+s+'"/>'
        :'<c r="'+ref+'" s="'+s+'"><v>'+sv+'</v></c>'; }
      else { cs+='<c r="'+ref+'" s="'+s+'" t="inlineStr"><is><t xml:space="preserve">'+esc(c.v)+'</t></is></c>'; }
    });
    sd+='<row r="'+r+'"'+(ln.h?' ht="'+ln.h+'" customHeight="1"':'')+'>'+cs+'</row>';
  });
  const ncols=Math.max(1,...linhas.map(l=>(l.cells||[]).length));
  const dim='A1:'+colName(ncols-1)+Math.max(1,linhas.length);
  let xml='<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'+
    '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">'+
    '<dimension ref="'+dim+'"/>'+
    '<sheetViews><sheetView workbookViewId="0" showGridLines="0"'+
      (cfg.ativa?' tabSelected="1"':'')+'>'+
      (congelar?'<pane ySplit="'+congelar+'" topLeftCell="A'+(congelar+1)+
        '" activePane="bottomLeft" state="frozen"/>':'')+
    '</sheetView></sheetViews>'+
    '<sheetFormatPr defaultRowHeight="15"/>';
  if(cols.length)xml+='<cols>'+cols.map((w,i)=>
    '<col min="'+(i+1)+'" max="'+(i+1)+'" width="'+w+'" customWidth="1"/>').join('')+'</cols>';
  xml+='<sheetData>'+sd+'</sheetData>';
  if(filtro)xml+='<autoFilter ref="'+filtro+'"/>';
  if(formatacao)xml+='<conditionalFormatting sqref="'+formatacao.ref+'">'+
    '<cfRule type="cellIs" dxfId="0" priority="1" operator="equal"><formula>"'+formatacao.igual+'"</formula></cfRule>'+
    '</conditionalFormatting>';
  if(validacao)xml+='<dataValidations count="1">'+
    '<dataValidation type="list" allowBlank="1" showInputMessage="1" showErrorMessage="1" '+
    'errorTitle="Valor inválido" error="Escolha uma das opções da lista." '+
    'promptTitle="Marcar" prompt="'+esc(validacao.dica||'')+'" sqref="'+validacao.ref+'">'+
    '<formula1>"'+validacao.opcoes.join(',')+'"</formula1></dataValidation></dataValidations>';
  xml+='<pageMargins left="0.4" right="0.4" top="0.5" bottom="0.5" header="0.3" footer="0.3"/>'+
    '<pageSetup orientation="landscape" fitToWidth="1" paperSize="9"/></worksheet>';
  return xml;
}

/* ---------------- pasta ---------------- */
function nomeAba(s){
  let n=String(s).replace(/[\\\/\?\*\[\]:]/g,'-').slice(0,31);
  return n||'Planilha';
}
function pasta(abas){
  const n=abas.length;
  const ct='<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'+
   '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'+
   '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'+
   '<Default Extension="xml" ContentType="application/xml"/>'+
   '<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>'+
   abas.map((a,i)=>'<Override PartName="/xl/worksheets/sheet'+(i+1)+
     '.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>').join('')+
   '<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>'+
   '<Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>'+
   '<Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/>'+
   '</Types>';
  const rels='<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'+
   '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'+
   '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>'+
   '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/>'+
   '<Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties" Target="docProps/app.xml"/>'+
   '</Relationships>';
  const wb='<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'+
   '<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" '+
   'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">'+
   '<sheets>'+abas.map((a,i)=>'<sheet name="'+esc(nomeAba(a.nome))+'" sheetId="'+(i+1)+
     '" r:id="rId'+(i+1)+'"/>').join('')+'</sheets></workbook>';
  const wbr='<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'+
   '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'+
   abas.map((a,i)=>'<Relationship Id="rId'+(i+1)+
     '" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet'+(i+1)+'.xml"/>').join('')+
   '<Relationship Id="rId'+(n+1)+'" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>'+
   '</Relationships>';
  const core='<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'+
   '<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" '+
   'xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" '+
   'xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">'+
   '<dc:title>Plano de Ação Bolsão 2027</dc:title><dc:creator>Matriz Educação</dc:creator>'+
   '<cp:lastModifiedBy>Matriz Educação</cp:lastModifiedBy></cp:coreProperties>';
  const app='<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'+
   '<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties" '+
   'xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes">'+
   '<Application>Matriz Educação</Application></Properties>';
  const arquivos=[
   {nome:'[Content_Types].xml',dados:ct},
   {nome:'_rels/.rels',dados:rels},
   {nome:'docProps/core.xml',dados:core},
   {nome:'docProps/app.xml',dados:app},
   {nome:'xl/workbook.xml',dados:wb},
   {nome:'xl/_rels/workbook.xml.rels',dados:wbr},
   {nome:'xl/styles.xml',dados:STYLES}];
  abas.forEach((a,i)=>arquivos.push({nome:'xl/worksheets/sheet'+(i+1)+'.xml',
    dados:folha(Object.assign({ativa:i===0},a))}));
  return zip(arquivos);
}
module.exports={pasta,colName,serial};
