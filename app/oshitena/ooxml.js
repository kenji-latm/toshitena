// オシテナ：押印欄（点線円）のOOXML生成
// Office.js の insertOoxml に渡すフラットOPC（pkg:package）を組み立てる。
// Word 2016以降・Mac・Web共通で動く WordApi 1.1 の範囲だけを使う。
//
// 1個の押印欄 = 円（楕円図形）＋ 円の下のキャプション（枠なしテキストボックス）。
// どちらも Word 標準の「wps/DrawingML＋VMLフォールバック」なので、
// 手作業で描いた図形と同じ扱いで選択・移動・削除できる。

(function (global) {
  "use strict";

  var EMU_PER_MM = 36000;
  var PT_PER_MM = 72 / 25.4;

  var CAPTION_FONT_PT = 8; // 円の下の文字サイズ
  var CAPTION_GAP_MM = 0.8; // 円とキャプションの間隔
  var CAPTION_HEIGHT_MM = 5; // キャプション枠の高さ
  var CAPTION_PAD_MM = 5; // キャプション枠を円より左右に広げる量（中央揃え用）

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

  var DASH_PRST = { solid: null, dash: "dash", sysDot: "sysDot" };
  var DASH_VML = { solid: null, dash: "dash", sysDot: "1 1" };

  function anchorOpen(offXmm, offYmm, wMm, hMm, relHeight) {
    return (
      '<wp:anchor distT="0" distB="0" distL="0" distR="0" simplePos="0" relativeHeight="' +
      relHeight +
      '" behindDoc="0" locked="0" layoutInCell="1" allowOverlap="1">' +
      '<wp:simplePos x="0" y="0"/>' +
      '<wp:positionH relativeFrom="character"><wp:posOffset>' + mmToEmu(offXmm) + "</wp:posOffset></wp:positionH>" +
      '<wp:positionV relativeFrom="line"><wp:posOffset>' + mmToEmu(offYmm) + "</wp:posOffset></wp:positionV>" +
      '<wp:extent cx="' + mmToEmu(wMm) + '" cy="' + mmToEmu(hMm) + '"/>'
    );
  }

  function vmlStyle(offXmm, offYmm, wMm, hMm, relHeight) {
    return (
      "position:absolute;margin-left:" + mmToPt(offXmm) + "pt;margin-top:" + mmToPt(offYmm) +
      "pt;width:" + mmToPt(wMm) + "pt;height:" + mmToPt(hMm) +
      "pt;z-index:" + relHeight +
      ";mso-position-horizontal-relative:char;mso-position-vertical-relative:line"
    );
  }

  /**
   * 円（楕円図形）1個分の <mc:AlternateContent> を作る。
   * @param {Object} opts
   * @param {number} opts.diameterMm  円の直径（mm）
   * @param {string} opts.color      線色（"7F7F7F" のような6桁HEX、#なし）
   * @param {string} opts.dash       "solid" | "dash" | "sysDot"
   * @param {number} opts.strokePt   線の太さ（pt）
   * @param {number} opts.offsetXmm  アンカー（カーソル位置）からの右方向オフセット（mm）
   * @param {number} opts.offsetYmm  同・下方向オフセット（mm）
   * @param {number} opts.docPrId    図形ID（文書内で一意の正整数）
   */
  function buildCircleXml(opts) {
    var d = opts.diameterMm;
    var strokeEmu = Math.round(opts.strokePt * 12700);
    var prstDash = DASH_PRST[opts.dash] || null;
    var vmlDash = DASH_VML[opts.dash] || null;
    var id = opts.docPrId;
    var relHeight = 251658240 + (id % 100000);

    var drawing =
      "<w:drawing>" +
      anchorOpen(opts.offsetXmm || 0, opts.offsetYmm || 0, d, d, relHeight) +
      '<wp:effectExtent l="' + strokeEmu + '" t="' + strokeEmu + '" r="' + strokeEmu + '" b="' + strokeEmu + '"/>' +
      "<wp:wrapNone/>" +
      '<wp:docPr id="' + id + '" name="押印欄' + id + '"/>' +
      "<wp:cNvGraphicFramePr/>" +
      '<a:graphic xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">' +
      '<a:graphicData uri="http://schemas.microsoft.com/office/word/2010/wordprocessingShape">' +
      "<wps:wsp>" +
      "<wps:cNvSpPr/>" +
      "<wps:spPr>" +
      '<a:xfrm><a:off x="0" y="0"/><a:ext cx="' + mmToEmu(d) + '" cy="' + mmToEmu(d) + '"/></a:xfrm>' +
      '<a:prstGeom prst="ellipse"><a:avLst/></a:prstGeom>' +
      "<a:noFill/>" +
      '<a:ln w="' + strokeEmu + '">' +
      '<a:solidFill><a:srgbClr val="' + opts.color + '"/></a:solidFill>' +
      (prstDash ? '<a:prstDash val="' + prstDash + '"/>' : "") +
      "</a:ln>" +
      "</wps:spPr>" +
      '<wps:bodyPr rot="0" spcFirstLastPara="0" vert="horz" wrap="square" lIns="0" tIns="0" rIns="0" bIns="0" numCol="1" anchor="ctr" anchorCtr="0"><a:noAutofit/></wps:bodyPr>' +
      "</wps:wsp>" +
      "</a:graphicData>" +
      "</a:graphic>" +
      "</wp:anchor>" +
      "</w:drawing>";

    var vml =
      "<w:pict>" +
      '<v:oval id="oshitena' + id + '" style="' +
      vmlStyle(opts.offsetXmm || 0, opts.offsetYmm || 0, d, d, relHeight) +
      '" filled="f" stroked="t" strokecolor="#' + opts.color +
      '" strokeweight="' + opts.strokePt + 'pt">' +
      (vmlDash ? '<v:stroke dashstyle="' + vmlDash + '"/>' : "") +
      "</v:oval>" +
      "</w:pict>";

    return (
      "<mc:AlternateContent>" +
      '<mc:Choice Requires="wps">' + drawing + "</mc:Choice>" +
      "<mc:Fallback>" + vml + "</mc:Fallback>" +
      "</mc:AlternateContent>"
    );
  }

  /**
   * 円の下のキャプション（枠なし・塗りなしテキストボックス）を作る。
   * 位置は円のオフセット＋直径から自動算出し、円の中心に対して左右中央になる。
   * @param {Object} opts buildCircleXml と同じ＋ caption（文字列）
   */
  function buildCaptionXml(opts) {
    var d = opts.diameterMm;
    var w = d + CAPTION_PAD_MM * 2;
    var h = CAPTION_HEIGHT_MM;
    var offX = (opts.offsetXmm || 0) - CAPTION_PAD_MM;
    var offY = (opts.offsetYmm || 0) + d + CAPTION_GAP_MM;
    var id = opts.docPrId;
    var relHeight = 251658240 + (id % 100000);
    var szHalfPt = Math.round(CAPTION_FONT_PT * 2);

    var runProps =
      "<w:rPr>" +
      '<w:rFonts w:ascii="游明朝" w:eastAsia="游明朝" w:hAnsi="游明朝"/>' +
      '<w:color w:val="' + opts.color + '"/>' +
      '<w:kern w:val="0"/>' +
      '<w:sz w:val="' + szHalfPt + '"/>' +
      '<w:szCs w:val="' + szHalfPt + '"/>' +
      "</w:rPr>";
    var captionPara =
      "<w:p>" +
      "<w:pPr>" +
      '<w:spacing w:before="0" w:after="0" w:line="240" w:lineRule="auto"/>' +
      '<w:jc w:val="center"/>' +
      runProps +
      "</w:pPr>" +
      "<w:r>" + runProps +
      '<w:t xml:space="preserve">' + escapeXml(opts.caption) + "</w:t>" +
      "</w:r>" +
      "</w:p>";

    var drawing =
      "<w:drawing>" +
      anchorOpen(offX, offY, w, h, relHeight) +
      '<wp:effectExtent l="0" t="0" r="0" b="0"/>' +
      "<wp:wrapNone/>" +
      '<wp:docPr id="' + id + '" name="押印欄ラベル' + id + '"/>' +
      "<wp:cNvGraphicFramePr/>" +
      '<a:graphic xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">' +
      '<a:graphicData uri="http://schemas.microsoft.com/office/word/2010/wordprocessingShape">' +
      "<wps:wsp>" +
      "<wps:cNvSpPr/>" +
      "<wps:spPr>" +
      '<a:xfrm><a:off x="0" y="0"/><a:ext cx="' + mmToEmu(w) + '" cy="' + mmToEmu(h) + '"/></a:xfrm>' +
      '<a:prstGeom prst="rect"><a:avLst/></a:prstGeom>' +
      "<a:noFill/>" +
      "<a:ln><a:noFill/></a:ln>" +
      "</wps:spPr>" +
      "<wps:txbx><w:txbxContent>" + captionPara + "</w:txbxContent></wps:txbx>" +
      '<wps:bodyPr rot="0" spcFirstLastPara="0" vert="horz" wrap="square" lIns="0" tIns="0" rIns="0" bIns="0" numCol="1" anchor="t" anchorCtr="0"><a:noAutofit/></wps:bodyPr>' +
      "</wps:wsp>" +
      "</a:graphicData>" +
      "</a:graphic>" +
      "</wp:anchor>" +
      "</w:drawing>";

    var vml =
      "<w:pict>" +
      '<v:rect id="oshitenacap' + id + '" style="' +
      vmlStyle(offX, offY, w, h, relHeight) +
      '" filled="f" stroked="f">' +
      '<v:textbox inset="0,0,0,0"><w:txbxContent>' + captionPara + "</w:txbxContent></v:textbox>" +
      "</v:rect>" +
      "</w:pict>";

    return (
      "<mc:AlternateContent>" +
      '<mc:Choice Requires="wps">' + drawing + "</mc:Choice>" +
      "<mc:Fallback>" + vml + "</mc:Fallback>" +
      "</mc:AlternateContent>"
    );
  }

  /** 円1個分（キャプション付きなら図形2個）のXMLを作る */
  function buildStampXml(opts) {
    var xml = buildCircleXml(opts);
    if (opts.caption) {
      xml += buildCaptionXml(
        Object.assign({}, opts, { docPrId: opts.docPrId + 1 })
      );
    }
    return xml;
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
   * shapesXml: buildStampXml の結果の配列
   */
  function buildDocumentXml(shapesXml) {
    return (
      "<w:document " + DOC_NS + ">" +
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
   * @param {Object} opts buildCircleXml と同じ（docPrId除く）＋ caption
   * @param {number} count 横に並べる個数
   * @param {number} gapMm 円同士の間隔（mm）
   */
  function buildStampOoxml(opts, count, gapMm) {
    var n = Math.max(1, count || 1);
    var gap = typeof gapMm === "number" ? gapMm : 4;
    var baseId = (Date.now() % 100000000) + 1000;
    var shapes = [];
    for (var i = 0; i < n; i++) {
      shapes.push(
        buildStampXml({
          diameterMm: opts.diameterMm,
          color: opts.color,
          dash: opts.dash,
          strokePt: opts.strokePt,
          caption: opts.caption || "",
          offsetXmm: (opts.offsetXmm || 0) + i * (opts.diameterMm + gap),
          offsetYmm: opts.offsetYmm || 0,
          docPrId: baseId + i * 2,
        })
      );
    }
    return buildPackage(shapes);
  }

  var api = {
    mmToEmu: mmToEmu,
    mmToPt: mmToPt,
    escapeXml: escapeXml,
    CAPTION_FONT_PT: CAPTION_FONT_PT,
    CAPTION_GAP_MM: CAPTION_GAP_MM,
    buildCircleXml: buildCircleXml,
    buildCaptionXml: buildCaptionXml,
    buildStampXml: buildStampXml,
    buildDocumentXml: buildDocumentXml,
    buildPackage: buildPackage,
    buildStampOoxml: buildStampOoxml,
  };

  if (typeof module !== "undefined" && module.exports) {
    module.exports = api;
  }
  global.OshitenaOoxml = api;
})(typeof window !== "undefined" ? window : globalThis);
