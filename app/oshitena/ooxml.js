// オシテナ：押印欄（点線円）のOOXML生成
// Office.js の insertOoxml に渡すフラットOPC（pkg:package）を組み立てる。
// Word 2016以降・Mac・Web共通で動く WordApi 1.1 の範囲だけを使う。
//
// 生成する図形は Word 標準の「楕円（wps/DrawingML）＋VMLフォールバック」なので、
// 手作業で描いた点線円と同じ扱いで選択・移動・削除できる。

(function (global) {
  "use strict";

  var EMU_PER_MM = 36000;
  var PT_PER_MM = 72 / 25.4;

  function mmToEmu(mm) {
    return Math.round(mm * EMU_PER_MM);
  }

  function mmToPt(mm) {
    return Math.round(mm * PT_PER_MM * 100) / 100;
  }

  function escapeXml(text) {
    return String(text)
      .replace(/&/g, "&amp;")
      .replace(/</g, "&lt;")
      .replace(/>/g, "&gt;")
      .replace(/"/g, "&quot;")
      .replace(/'/g, "&apos;");
  }

  // ラベル文字数に応じて円に収まるフォントサイズ（pt）を決める
  function labelFontPt(diameterMm, label) {
    var chars = Array.from(label).length;
    var base = diameterMm * PT_PER_MM * 0.44;
    var factor = chars <= 1 ? 1 : chars === 2 ? 0.62 : chars === 3 ? 0.44 : 0.34;
    return Math.max(6, Math.round(base * factor));
  }

  var DASH_PRST = { solid: null, dash: "dash", sysDot: "sysDot" };
  var DASH_VML = { solid: null, dash: "dash", sysDot: "1 1" };

  /**
   * 押印欄1個分の <mc:AlternateContent>（DrawingML＋VMLフォールバック）を作る。
   * @param {Object} opts
   * @param {number} opts.diameterMm  円の直径（mm）
   * @param {string} opts.color      線・文字色（"D93A22" のような6桁HEX、#なし）
   * @param {string} opts.dash       "solid" | "dash" | "sysDot"
   * @param {number} opts.strokePt   線の太さ（pt）
   * @param {string} opts.label      円内の文字（空文字ならなし）
   * @param {number} opts.offsetXmm  アンカー（カーソル位置）からの右方向オフセット（mm）
   * @param {number} opts.offsetYmm  同・下方向オフセット（mm）
   * @param {number} opts.docPrId    図形ID（文書内で一意の正整数）
   */
  function buildShapeXml(opts) {
    var cx = mmToEmu(opts.diameterMm);
    var cy = cx;
    var offX = mmToEmu(opts.offsetXmm || 0);
    var offY = mmToEmu(opts.offsetYmm || 0);
    var strokeEmu = Math.round(opts.strokePt * 12700);
    var color = opts.color;
    var prstDash = DASH_PRST[opts.dash] || null;
    var vmlDash = DASH_VML[opts.dash] || null;
    var id = opts.docPrId;
    var relHeight = 251658240 + (id % 1000);

    var txbx = "";
    var vmlTxbx = "";
    if (opts.label) {
      var fontPt = labelFontPt(opts.diameterMm, opts.label);
      var szHalfPt = Math.round(fontPt * 2);
      var runProps =
        '<w:rPr>' +
        '<w:rFonts w:ascii="游明朝" w:eastAsia="游明朝" w:hAnsi="游明朝"/>' +
        '<w:color w:val="' + color + '"/>' +
        '<w:kern w:val="0"/>' +
        '<w:sz w:val="' + szHalfPt + '"/>' +
        '<w:szCs w:val="' + szHalfPt + '"/>' +
        "</w:rPr>";
      var labelPara =
        "<w:p>" +
        "<w:pPr>" +
        '<w:spacing w:before="0" w:after="0" w:line="240" w:lineRule="auto"/>' +
        '<w:jc w:val="center"/>' +
        runProps +
        "</w:pPr>" +
        "<w:r>" + runProps +
        '<w:t xml:space="preserve">' + escapeXml(opts.label) + "</w:t>" +
        "</w:r>" +
        "</w:p>";
      txbx =
        "<wps:txbx><w:txbxContent>" + labelPara + "</w:txbxContent></wps:txbx>";
      vmlTxbx =
        '<v:textbox inset="0,0,0,0"><w:txbxContent>' +
        labelPara +
        "</w:txbxContent></v:textbox>";
    }

    var drawing =
      "<w:drawing>" +
      '<wp:anchor distT="0" distB="0" distL="0" distR="0" simplePos="0" relativeHeight="' +
      relHeight +
      '" behindDoc="0" locked="0" layoutInCell="1" allowOverlap="1">' +
      '<wp:simplePos x="0" y="0"/>' +
      '<wp:positionH relativeFrom="character"><wp:posOffset>' + offX + "</wp:posOffset></wp:positionH>" +
      '<wp:positionV relativeFrom="line"><wp:posOffset>' + offY + "</wp:posOffset></wp:positionV>" +
      '<wp:extent cx="' + cx + '" cy="' + cy + '"/>' +
      '<wp:effectExtent l="' + strokeEmu + '" t="' + strokeEmu + '" r="' + strokeEmu + '" b="' + strokeEmu + '"/>' +
      "<wp:wrapNone/>" +
      '<wp:docPr id="' + id + '" name="押印欄' + id + '"/>' +
      "<wp:cNvGraphicFramePr/>" +
      '<a:graphic xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">' +
      '<a:graphicData uri="http://schemas.microsoft.com/office/word/2010/wordprocessingShape">' +
      "<wps:wsp>" +
      "<wps:cNvSpPr/>" +
      "<wps:spPr>" +
      '<a:xfrm><a:off x="0" y="0"/><a:ext cx="' + cx + '" cy="' + cy + '"/></a:xfrm>' +
      '<a:prstGeom prst="ellipse"><a:avLst/></a:prstGeom>' +
      "<a:noFill/>" +
      '<a:ln w="' + strokeEmu + '">' +
      '<a:solidFill><a:srgbClr val="' + color + '"/></a:solidFill>' +
      (prstDash ? '<a:prstDash val="' + prstDash + '"/>' : "") +
      "</a:ln>" +
      "</wps:spPr>" +
      txbx +
      '<wps:bodyPr rot="0" spcFirstLastPara="0" vert="horz" wrap="square" lIns="0" tIns="0" rIns="0" bIns="0" numCol="1" anchor="ctr" anchorCtr="0"><a:noAutofit/></wps:bodyPr>' +
      "</wps:wsp>" +
      "</a:graphicData>" +
      "</a:graphic>" +
      "</wp:anchor>" +
      "</w:drawing>";

    var wPt = mmToPt(opts.diameterMm);
    var vml =
      "<w:pict>" +
      '<v:oval id="oshitena' + id + '" style="position:absolute;margin-left:' +
      mmToPt(opts.offsetXmm || 0) + "pt;margin-top:" + mmToPt(opts.offsetYmm || 0) +
      "pt;width:" + wPt + "pt;height:" + wPt +
      "pt;z-index:" + relHeight +
      ';mso-position-horizontal-relative:char;mso-position-vertical-relative:line" filled="f" stroked="t" strokecolor="#' +
      color + '" strokeweight="' + opts.strokePt + 'pt">' +
      (vmlDash ? '<v:stroke dashstyle="' + vmlDash + '"/>' : "") +
      vmlTxbx +
      "</v:oval>" +
      "</w:pict>";

    return (
      "<mc:AlternateContent>" +
      '<mc:Choice Requires="wps">' + drawing + "</mc:Choice>" +
      "<mc:Fallback>" + vml + "</mc:Fallback>" +
      "</mc:AlternateContent>"
    );
  }

  var DOC_NS =
    'xmlns:wpc="http://schemas.microsoft.com/office/word/2010/wordprocessingCanvas" ' +
    'xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" ' +
    'xmlns:o="urn:schemas-microsoft-com:office:office" ' +
    'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" ' +
    'xmlns:v="urn:schemas-microsoft-com:vml" ' +
    'xmlns:wp14="http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing" ' +
    'xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" ' +
    'xmlns:w10="urn:schemas-microsoft-com:office:word" ' +
    'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" ' +
    'xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml" ' +
    'xmlns:wpg="http://schemas.microsoft.com/office/word/2010/wordprocessingGroup" ' +
    'xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape" ' +
    'mc:Ignorable="w14 wp14"';

  /**
   * 図形入りの段落1つを持つ document.xml の中身を作る（フラットOPCの部品）。
   * shapes: buildShapeXml の結果の配列（複数個を同じ段落に入れる）
   */
  function buildDocumentXml(shapesXml) {
    return (
      '<w:document ' + DOC_NS + ">" +
      "<w:body>" +
      "<w:p><w:r>" + shapesXml.join("") + "</w:r></w:p>" +
      "</w:body>" +
      "</w:document>"
    );
  }

  /** insertOoxml に渡すフラットOPCパッケージ全体を作る */
  function buildPackage(shapesXml) {
    return (
      '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
      '<pkg:package xmlns:pkg="http://schemas.microsoft.com/office/2006/xmlPackage">' +
      '<pkg:part pkg:name="/_rels/.rels" pkg:contentType="application/vnd.openxmlformats-package.relationships+xml" pkg:padding="512">' +
      "<pkg:xmlData>" +
      '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
      '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>' +
      "</Relationships>" +
      "</pkg:xmlData>" +
      "</pkg:part>" +
      '<pkg:part pkg:name="/word/document.xml" pkg:contentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml">' +
      "<pkg:xmlData>" +
      buildDocumentXml(shapesXml) +
      "</pkg:xmlData>" +
      "</pkg:part>" +
      "</pkg:package>"
    );
  }

  /**
   * 挿入用OOXMLを組み立てる入口。
   * @param {Object} opts buildShapeXml と同じ（docPrId除く）
   * @param {number} count 横に並べる個数（捨印欄などで使用）
   * @param {number} gapMm 円同士の間隔（mm）
   */
  function buildStampOoxml(opts, count, gapMm) {
    var n = Math.max(1, count || 1);
    var gap = typeof gapMm === "number" ? gapMm : 4;
    var baseId = (Date.now() % 100000000) + 1000;
    var shapes = [];
    for (var i = 0; i < n; i++) {
      shapes.push(
        buildShapeXml({
          diameterMm: opts.diameterMm,
          color: opts.color,
          dash: opts.dash,
          strokePt: opts.strokePt,
          label: opts.label,
          offsetXmm: (opts.offsetXmm || 0) + i * (opts.diameterMm + gap),
          offsetYmm: opts.offsetYmm || 0,
          docPrId: baseId + i,
        })
      );
    }
    return buildPackage(shapes);
  }

  var api = {
    mmToEmu: mmToEmu,
    mmToPt: mmToPt,
    labelFontPt: labelFontPt,
    escapeXml: escapeXml,
    buildShapeXml: buildShapeXml,
    buildDocumentXml: buildDocumentXml,
    buildPackage: buildPackage,
    buildStampOoxml: buildStampOoxml,
  };

  if (typeof module !== "undefined" && module.exports) {
    module.exports = api;
  }
  global.OshitenaOoxml = api;
})(typeof window !== "undefined" ? window : globalThis);
