"""Small deterministic OOXML fixture, not a WPS-client export."""
import struct
import zipfile
import zlib


def png(width, height):
    def chunk(kind, data):
        return struct.pack('>I', len(data)) + kind + data + struct.pack('>I', zlib.crc32(kind + data))
    return (b'\x89PNG\r\n\x1a\n'
            + chunk(b'IHDR', struct.pack('>IIBBBBB', width, height, 8, 2, 0, 0, 0))
            + chunk(b'IDAT', zlib.compress((b'\0' + b'\x28\x80\xc0' * width) * height))
            + chunk(b'IEND', b''))


def build(path, *, missing=False, formula=None):
    main = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'
    rel = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
    pr = 'http://schemas.openxmlformats.org/package/2006/relationships'
    xdr = 'http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing'
    a = 'http://schemas.openxmlformats.org/drawingml/2006/main'
    formulas = [formula or '_xlfn.DISPIMG("wide",1)', 'dispimg ( "tall" , 1 )', 'DISPIMG("wide",1)']
    cells = ['A1', 'B2', 'D5']
    rows = ''.join(f'<row r="{cell[1:]}" ht="60" customHeight="1"><c r="{cell}" t="str"><f>{f}</f><v>cached</v></c></row>' for cell, f in zip(cells, formulas))
    pics = ''.join(f'<etc:cellImage><xdr:pic><xdr:nvPicPr><xdr:cNvPr id="{n}" name="{name}"/></xdr:nvPicPr><xdr:blipFill><a:blip r:embed="rId{n}"/></xdr:blipFill></xdr:pic></etc:cellImage>' for n, name in enumerate(['wide', 'tall'], 1))
    parts = {
        '[Content_Types].xml': '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Default Extension="png" ContentType="image/png"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/><Override PartName="/xl/cellimages.xml" ContentType="application/vnd.wps-officedocument.cellimage+xml"/></Types>',
        '_rels/.rels': f'<Relationships xmlns="{pr}"><Relationship Id="rId1" Type="{rel}/officeDocument" Target="xl/workbook.xml"/></Relationships>',
        'xl/workbook.xml': f'<workbook xmlns="{main}" xmlns:r="{rel}"><sheets><sheet name="Sheet1" sheetId="1" r:id="rId1"/></sheets></workbook>',
        'xl/_rels/workbook.xml.rels': f'<Relationships xmlns="{pr}"><Relationship Id="rId1" Type="{rel}/worksheet" Target="/xl/worksheets/sheet1.xml"/><Relationship Id="rId2" Type="http://www.wps.cn/officeDocument/2020/cellImage" Target="cellimages.xml"/></Relationships>',
        'xl/worksheets/sheet1.xml': f'<worksheet xmlns="{main}"><sheetFormatPr defaultRowHeight="15" defaultColWidth="8.43"/><cols><col min="2" max="3" width="20" customWidth="1"/></cols><sheetData>{rows}</sheetData><mergeCells count="1"><mergeCell ref="B2:C3"/></mergeCells></worksheet>',
        'xl/cellimages.xml': f'<etc:cellImages xmlns:etc="http://www.wps.cn/officeDocument/2017/etCustomData" xmlns:xdr="{xdr}" xmlns:a="{a}" xmlns:r="{rel}">{pics}</etc:cellImages>',
        'xl/_rels/cellimages.xml.rels': f'<Relationships xmlns="{pr}"><Relationship Id="rId1" Type="{rel}/image" Target="media/wide.png"/><Relationship Id="rId2" Type="{rel}/image" Target="media/tall.png"/></Relationships>',
        'xl/media/wide.png': png(120, 30),
        'xl/media/tall.png': png(30, 120),
        'custom/preserve.bin': b'preserve unrelated package part',
    }
    if missing:
        del parts['xl/media/tall.png']
    with zipfile.ZipFile(path, 'w', zipfile.ZIP_DEFLATED) as archive:
        for name, data in parts.items():
            archive.writestr(name, data)
    return path
