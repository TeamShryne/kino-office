package com.teamshryne.kinooffice

import android.content.Context
import org.apache.poi.xwpf.usermodel.*
import org.apache.poi.xssf.usermodel.*
import org.apache.poi.xslf.usermodel.*
import org.apache.poi.ss.usermodel.*
import org.apache.poi.hssf.usermodel.HSSFWorkbook
import org.json.JSONArray
import org.json.JSONObject
import java.io.File
import java.io.FileInputStream
import java.io.FileOutputStream
import java.util.Locale
import java.util.UUID

data class OfficeLimits(
    val maxSheets: Int = 20,
    val maxRows: Int = 5000,
    val maxCells: Int = 100000,
    val maxSlides: Int = 50,
    val maxParagraphs: Int = 500,
    val maxChars: Int = 200000,
    val maxBytes: Long = 25L * 1024 * 1024,
)

class OfficeRenderer(
    private val directory: File,
    private val limits: OfficeLimits = OfficeLimits(),
) {
    init { directory.mkdirs() }

    fun render(arguments: org.json.JSONObject): String {
        val format = arguments.getString("format")
        val operation = arguments.optString("operation", "create")
        val title = arguments.optString("title", "Kino document")
        val outputFilename = arguments.optString("output_filename").takeIf { it.isNotBlank() }
        return if (operation == "edit") {
            val filename = arguments.getString("filename")
            val operations = arguments.optJSONArray("operations") ?: org.json.JSONArray()
            edit(format, filename, operations, outputFilename)
        } else {
            create(format, title, arguments, outputFilename)
        }
    }

    fun create(format: String, title: String, payload: JSONObject, outputFilename: String? = null): String {
        val safeTitle = title.take(120).ifBlank { "Kino document" }
        val name = safeOutputName(outputFilename, format) ?: "kino-${UUID.randomUUID()}.$format"
        val output = File(directory, name)
        try {
            when (format) {
                "docx" -> createDocx(output, safeTitle, payload)
                "xlsx" -> createXlsx(output, safeTitle, payload, false)
                "xls" -> createXlsx(output, safeTitle, payload, true)
                "pptx" -> createPptx(output, safeTitle, payload)
                else -> error("Unsupported office format: $format")
            }
            check(output.isFile && output.length() in 1..limits.maxBytes) { "Office output is empty or too large." }
            return output.name
        } catch (error: Throwable) {
            output.delete()
            throw error
        }
    }

    fun edit(format: String, filename: String, operations: JSONArray, outputFilename: String? = null): String {
        require(filename.matches(Regex("kino-[a-f0-9-]{36}\\.(docx|xlsx|pptx|xls)"))) { "Invalid office filename." }
        val source = File(directory, filename)
        require(source.isFile) { "Office file not found." }
        val ext = filename.substringAfterLast('.')
        val name = safeOutputName(outputFilename, ext) ?: "kino-${UUID.randomUUID()}.$ext"
        val output = File(directory, name)
        try {
            when (format) {
                "docx" -> editDocx(source, output, operations)
                "xlsx", "xls" -> editXlsx(source, output, operations, format == "xls")
                "pptx" -> editPptx(source, output, operations)
                else -> error("Unsupported office format: $format")
            }
            check(output.isFile && output.length() in 1..limits.maxBytes) { "Office output is empty or too large." }
            return output.name
        } catch (error: Throwable) {
            output.delete()
            throw error
        }
    }

    // ---------- DOCX ----------
    private fun createDocx(output: File, title: String, payload: JSONObject) {
        val doc = XWPFDocument()
        val styles = payload.optJSONObject("style")
        val baseFont = styles?.optString("font").orEmpty().ifBlank { "Calibri" }
        val baseSize = styles?.optInt("base_size", 11) ?: 11
        val normalStyle = doc.createParagraph().style
        // Title
        if (title.isNotBlank()) {
            val p = doc.createParagraph()
            p.alignment = ParagraphAlignment.LEFT
            val r = p.createRun()
            r.setText(title)
            r.isBold = true
            r.fontSize = baseSize + 6
            r.fontFamily = baseFont
            r.setColor("1F3864")
        }
        if (payload.has("blocks")) {
            renderDocxBlocks(doc, payload.optJSONArray("blocks")!!, baseFont, baseSize)
        } else {
            val markdown = payload.optString("markdown", "")
            renderMarkdownToDocx(doc, markdown, baseFont, baseSize)
        }
        FileOutputStream(output).use { doc.write(it) }
        doc.close()
    }

    private fun renderDocxBlocks(doc: XWPFDocument, blocks: JSONArray, baseFont: String, baseSize: Int) {
        var count = 0
        for (i in 0 until blocks.length()) {
            if (count++ > limits.maxParagraphs) break
            val b = blocks.optJSONObject(i) ?: continue
            when (b.optString("type")) {
                "heading" -> {
                    val level = b.optInt("level", 1).coerceIn(1, 4)
                    val p = doc.createParagraph()
                    p.style = "Heading$level"
                    val r = p.createRun()
                    r.setText(b.optString("text"))
                    r.isBold = true
                    r.fontFamily = baseFont
                    r.fontSize = baseSize + (5 - level)
                }
                "paragraph" -> {
                    val p = doc.createParagraph()
                    val r = p.createRun()
                    r.setText(b.optString("text"))
                    r.fontFamily = baseFont
                    r.fontSize = baseSize
                    if (b.optBoolean("bold")) r.isBold = true
                    if (b.optBoolean("italic")) r.isItalic = true
                    val color = b.optString("color").trim()
                    if (color.startsWith("#")) r.setColor(color.removePrefix("#"))
                }
                "bullet" -> {
                    val p = doc.createParagraph()
                    p.numID = null
                    val r = p.createRun()
                    r.setText("• " + b.optString("text"))
                    r.fontFamily = baseFont
                    r.fontSize = baseSize
                }
                "table" -> {
                    val rows = b.optJSONArray("rows") ?: continue
                    if (rows.length() == 0) continue
                    val first = rows.optJSONArray(0) ?: continue
                    val table = doc.createTable(rows.length(), first.length())
                    for (ri in 0 until rows.length()) {
                        val row = rows.optJSONArray(ri) ?: continue
                        val cells = table.getRow(ri).tableCells
                        for (ci in 0 until row.length()) {
                            if (ci >= cells.size) break
                            cells[ci].setText(row.opt(ci)?.toString().orEmpty())
                        }
                    }
                }
                "page_break" -> doc.createParagraph().createRun().addBreak(BreakType.PAGE)
            }
        }
    }

    private fun renderMarkdownToDocx(doc: XWPFDocument, markdown: String, baseFont: String, baseSize: Int) {
        val lines = markdown.lines()
        var i = 0
        var count = 0
        while (i < lines.size && count < limits.maxParagraphs) {
            val line = lines[i]
            val trimmed = line.trim()
            if (trimmed.isEmpty()) { i++; continue }
            when {
                trimmed.startsWith("# ") -> {
                    val p = doc.createParagraph(); p.style = "Heading1"
                    val r = p.createRun(); r.setText(trimmed.removePrefix("# ").trim()); r.isBold = true; r.fontSize = baseSize + 6; r.fontFamily = baseFont
                }
                trimmed.startsWith("## ") -> {
                    val p = doc.createParagraph(); p.style = "Heading2"
                    val r = p.createRun(); r.setText(trimmed.removePrefix("## ").trim()); r.isBold = true; r.fontSize = baseSize + 4; r.fontFamily = baseFont
                }
                trimmed.startsWith("### ") -> {
                    val p = doc.createParagraph(); p.style = "Heading3"
                    val r = p.createRun(); r.setText(trimmed.removePrefix("### ").trim()); r.isBold = true; r.fontSize = baseSize + 2; r.fontFamily = baseFont
                }
                trimmed.startsWith("- ") || trimmed.startsWith("* ") -> {
                    val p = doc.createParagraph()
                    val r = p.createRun(); r.setText("• " + trimmed.drop(2).trim()); r.fontFamily = baseFont; r.fontSize = baseSize
                }
                trimmed.startsWith("|") && i + 1 < lines.size && lines[i + 1].trim().startsWith("|") -> {
                    // table
                    val tableLines = mutableListOf<String>()
                    var j = i
                    while (j < lines.size && lines[j].trim().startsWith("|")) {
                        tableLines += lines[j].trim()
                        j++
                    }
                    val rows = tableLines.filterNot { it.matches(Regex("^\\|[\\s|:-]+\\|$")) }
                    if (rows.isNotEmpty()) {
                        val cols = rows.first().trim('|').split("|").size
                        val table = doc.createTable(rows.size, cols)
                        rows.forEachIndexed { ri, rl ->
                            val cells = table.getRow(ri).tableCells
                            val parts = rl.trim('|').split("|")
                            parts.forEachIndexed { ci, value ->
                                if (ci < cells.size) cells[ci].setText(value.trim())
                            }
                        }
                    }
                    i = j - 1
                }
                else -> {
                    val p = doc.createParagraph()
                    val r = p.createRun(); r.setText(trimmed); r.fontFamily = baseFont; r.fontSize = baseSize
                }
            }
            count++
            i++
        }
    }

    private fun editDocx(source: File, output: File, operations: JSONArray) {
        val doc = XWPFDocument(FileInputStream(source))
        for (i in 0 until operations.length()) {
            val op = operations.optJSONObject(i) ?: continue
            when (op.optString("type")) {
                "append_paragraph" -> {
                    val p = doc.createParagraph()
                    val r = p.createRun()
                    r.setText(op.optString("text"))
                }
                "append_heading" -> {
                    val p = doc.createParagraph()
                    p.style = "Heading${op.optInt("level", 1).coerceIn(1, 4)}"
                    val r = p.createRun()
                    r.setText(op.optString("text"))
                    r.isBold = true
                }
                "replace_text" -> {
                    val find = op.optString("find")
                    val replace = op.optString("replace")
                    if (find.isBlank()) continue
                    for (p in doc.paragraphs) {
                        for (r in p.runs) {
                            val t = r.text()
                            if (t != null && t.contains(find)) r.setText(t.replace(find, replace), 0)
                        }
                    }
                }
                "append_table" -> {
                    val rows = op.optJSONArray("rows") ?: continue
                    if (rows.length() == 0) continue
                    val first = rows.optJSONArray(0) ?: continue
                    val table = doc.createTable(rows.length(), first.length())
                    for (ri in 0 until rows.length()) {
                        val row = rows.optJSONArray(ri) ?: continue
                        val cells = table.getRow(ri).tableCells
                        for (ci in 0 until row.length()) {
                            if (ci < cells.size) cells[ci].setText(row.opt(ci)?.toString().orEmpty())
                        }
                    }
                }
            }
        }
        FileOutputStream(output).use { doc.write(it) }
        doc.close()
    }

    // ---------- XLSX ----------
    private fun createXlsx(output: File, title: String, payload: JSONObject, legacy: Boolean) {
        val wb: Workbook = if (legacy) HSSFWorkbook() else XSSFWorkbook()
        val sheets = payload.optJSONArray("sheets")
        if (sheets == null || sheets.length() == 0) {
            val sh = wb.createSheet("Sheet1")
            val r = sh.createRow(0)
            r.createCell(0).setCellValue(title)
        } else {
            require(sheets.length() <= limits.maxSheets) { "Too many sheets." }
            for (si in 0 until sheets.length()) {
                val s = sheets.optJSONObject(si) ?: continue
                val name = sanitizeSheetName(s.optString("name").ifBlank { "Sheet${si + 1}" }, si)
                val sh = wb.createSheet(name)
                val rows = if (s.has("rows")) s.optJSONArray("rows") else if (s.has("csv")) parseCsv(s.optString("csv")) else null
                val maxRows = minOf(rows?.length() ?: 0, limits.maxRows)
                var cellCount = 0
                for (ri in 0 until maxRows) {
                    val row = rows?.optJSONArray(ri) ?: continue
                    val r = sh.createRow(ri)
                    for (ci in 0 until row.length()) {
                        if (cellCount++ > limits.maxCells) break
                        val cell = r.createCell(ci)
                        setCellValue(cell, row.opt(ci))
                    }
                }
                // column widths
                val widths = s.optJSONArray("column_widths")
                if (widths != null) for (ci in 0 until widths.length()) {
                    val w = widths.optInt(ci, -1)
                    if (w > 0) sh.setColumnWidth(ci, w * 256)
                }
                val freeze = s.optString("freeze_panes")
                if (freeze.isNotBlank()) {
                    val col = freeze.takeWhile { it.isLetter() }.uppercase().let { colLettersToIndex(it) }
                    val row = freeze.dropWhile { it.isLetter() }.toIntOrNull()?.minus(1) ?: 0
                    sh.createFreezePane(col, row.coerceAtLeast(0))
                }
                // basic formats
                val formats = s.optJSONObject("formats")
                if (formats != null) {
                    val keys = formats.keys()
                    while (keys.hasNext()) {
                        val range = keys.next()
                        val fmt = formats.optJSONObject(range) ?: continue
                        applyCellFormat(wb, sh, range, fmt)
                    }
                }
            }
        }
        FileOutputStream(output).use { wb.write(it) }
        wb.close()
    }

    private fun editXlsx(source: File, output: File, operations: JSONArray, legacy: Boolean) {
        val wb: Workbook = if (legacy) HSSFWorkbook(FileInputStream(source)) else XSSFWorkbook(FileInputStream(source))
        for (i in 0 until operations.length()) {
            val op = operations.optJSONObject(i) ?: continue
            when (op.optString("type")) {
                "append_row" -> {
                    val sheetName = op.optString("sheet")
                    val sh = if (sheetName.isBlank()) wb.getSheetAt(0) else wb.getSheet(sheetName) ?: wb.getSheetAt(0)
                    val row = sh.createRow(sh.lastRowNum + 1)
                    val values = op.optJSONArray("values") ?: continue
                    for (ci in 0 until values.length()) setCellValue(row.createCell(ci), values.opt(ci))
                }
                "set_cell" -> {
                    val sheetName = op.optString("sheet")
                    val sh = if (sheetName.isBlank()) wb.getSheetAt(0) else wb.getSheet(sheetName) ?: wb.getSheetAt(0)
                    val cellRef = op.optString("cell")
                    val (col, row) = parseCellRef(cellRef)
                    val r = sh.getRow(row) ?: sh.createRow(row)
                    val c = r.getCell(col) ?: r.createCell(col)
                    setCellValue(c, op.opt("value"))
                }
                "add_sheet" -> {
                    val name = sanitizeSheetName(op.optString("name").ifBlank { "Sheet${wb.numberOfSheets + 1}" }, wb.numberOfSheets)
                    wb.createSheet(name)
                }
                "rename_sheet" -> {
                    val oldName = op.optString("sheet")
                    val sh = if (oldName.isBlank()) wb.getSheetAt(0) else wb.getSheet(oldName) ?: wb.getSheetAt(0)
                    val idx = wb.getSheetIndex(sh)
                    wb.setSheetName(idx, sanitizeSheetName(op.optString("new_name", sh.sheetName), idx))
                }
                "set_column_width" -> {
                    val sheetName = op.optString("sheet")
                    val sh = if (sheetName.isBlank()) wb.getSheetAt(0) else wb.getSheet(sheetName) ?: wb.getSheetAt(0)
                    val col = op.optInt("column", 0)
                    val width = op.optInt("width", 12)
                    if (col >= 0) sh.setColumnWidth(col, width * 256)
                }
            }
        }
        FileOutputStream(output).use { wb.write(it) }
        wb.close()
    }

    // ---------- PPTX ----------
    private fun createPptx(output: File, title: String, payload: JSONObject) {
        val ppt = XMLSlideShow()
        val slides = payload.optJSONArray("slides")
        val master = ppt.slideMasters.firstOrNull()
        val layout = master?.getLayout(org.apache.poi.xslf.usermodel.SlideLayout.TITLE_AND_CONTENT)
            ?: master?.getLayout(org.apache.poi.xslf.usermodel.SlideLayout.TITLE_ONLY)
        if (slides == null || slides.length() == 0) {
            val slide = if (layout != null) ppt.createSlide(layout) else ppt.createSlide()
            val titleShape = slide.getPlaceholder(0) as? org.apache.poi.xslf.usermodel.XSLFTextShape
            titleShape?.let { it.clearText(); it.setText(title) }
        } else {
            require(slides.length() <= limits.maxSlides) { "Too many slides." }
            for (si in 0 until slides.length()) {
                val s = slides.optJSONObject(si) ?: continue
                val slide = if (layout != null) ppt.createSlide(layout) else ppt.createSlide()
                val titleText = s.optString("title").ifBlank { if (si == 0) title else "Slide ${si + 1}" }
                val titleShape = slide.getPlaceholder(0) as? org.apache.poi.xslf.usermodel.XSLFTextShape
                titleShape?.let { it.clearText(); it.setText(titleText) }
                val bullets = s.optJSONArray("bullets")
                if (bullets != null && bullets.length() > 0) {
                    val body = slide.getPlaceholder(1) as? org.apache.poi.xslf.usermodel.XSLFTextShape
                    body?.clearText()
                    for (bi in 0 until bullets.length()) {
                        val bp = body?.addNewTextParagraph() ?: continue
                        bp.bulletCharacter = "•"
                        val br = bp.addNewTextRun()
                        br.setText(bullets.optString(bi))
                        br.fontSize = 18.0
                    }
                }
                val notes = s.optString("notes")
                if (notes.isNotBlank()) {
                    val notesSlide = slide.notes
                    val notesShape = notesSlide?.getPlaceholder(0) as? org.apache.poi.xslf.usermodel.XSLFTextShape
                    notesShape?.setText(notes)
                }
            }
        }
        FileOutputStream(output).use { ppt.write(it) }
        ppt.close()
    }

    private fun editPptx(source: File, output: File, operations: JSONArray) {
        val ppt = XMLSlideShow(FileInputStream(source))
        val master = ppt.slideMasters.firstOrNull()
        val layout = master?.getLayout(org.apache.poi.xslf.usermodel.SlideLayout.TITLE_AND_CONTENT)
            ?: master?.getLayout(org.apache.poi.xslf.usermodel.SlideLayout.TITLE_ONLY)
        for (i in 0 until operations.length()) {
            val op = operations.optJSONObject(i) ?: continue
            when (op.optString("type")) {
                "add_slide" -> {
                    val slide = if (layout != null) ppt.createSlide(layout) else ppt.createSlide()
                    val title = op.optString("title")
                    if (title.isNotBlank()) {
                        val ts = slide.getPlaceholder(0) as? org.apache.poi.xslf.usermodel.XSLFTextShape
                        ts?.let { it.clearText(); it.setText(title) }
                    }
                    val bullets = op.optJSONArray("bullets")
                    if (bullets != null && bullets.length() > 0) {
                        val body = slide.getPlaceholder(1) as? org.apache.poi.xslf.usermodel.XSLFTextShape
                        body?.clearText()
                        for (bi in 0 until bullets.length()) {
                            val bp = body?.addNewTextParagraph() ?: continue
                            bp.bulletCharacter = "•"
                            bp.addNewTextRun().setText(bullets.optString(bi))
                        }
                    }
                }
                "append_bullet" -> {
                    val index = op.optInt("slide", ppt.slides.size - 1)
                    if (index < 0 || index >= ppt.slides.size) continue
                    val slide = ppt.slides[index]
                    val body = slide.getPlaceholder(1) as? org.apache.poi.xslf.usermodel.XSLFTextShape
                    val bp = body?.addNewTextParagraph() ?: continue
                    bp.bulletCharacter = "•"
                    bp.addNewTextRun().setText(op.optString("text"))
                }
                "set_title" -> {
                    val index = op.optInt("slide", 0)
                    if (index < 0 || index >= ppt.slides.size) continue
                    val slide = ppt.slides[index]
                    val first = slide.getPlaceholder(0) as? org.apache.poi.xslf.usermodel.XSLFTextShape
                    first?.let { it.clearText(); it.setText(op.optString("title")) }
                }
                "delete_slide" -> {
                    val index = op.optInt("slide", -1)
                    if (index in 0 until ppt.slides.size) ppt.removeSlide(index)
                }
            }
        }
        FileOutputStream(output).use { ppt.write(it) }
        ppt.close()
    }

    // ---------- helpers ----------
    private fun safeOutputName(name: String?, format: String): String? {
        val trimmed = name?.trim().orEmpty()
        if (trimmed.isEmpty()) return null
        val base = trimmed.substringBeforeLast('.').ifBlank { trimmed }
        val safe = base.replace(Regex("[^a-zA-Z0-9._-]"), "_").take(80)
        return if (safe.isBlank()) null else "$safe.$format"
    }

    private fun sanitizeSheetName(name: String, index: Int): String {
        var n = name.replace(Regex("[\\[\\]:*?/\\\\]"), "_").take(31)
        if (n.isBlank()) n = "Sheet${index + 1}"
        return n
    }

    private fun setCellValue(cell: Cell, value: Any?) {
        when (value) {
            null -> cell.setBlank()
            is Boolean -> cell.setCellValue(value)
            is Number -> cell.setCellValue(value.toDouble())
            else -> cell.setCellValue(value.toString())
        }
    }

    private fun parseCsv(csv: String): JSONArray {
        val rows = JSONArray()
        csv.lines().filter { it.isNotBlank() }.forEach { line ->
            val cells = JSONArray()
            // simple CSV split
            line.split(",").forEach { cells.put(it.trim().removeSurrounding("\"")) }
            rows.put(cells)
        }
        return rows
    }

    private fun colLettersToIndex(letters: String): Int {
        var result = 0
        for (c in letters) if (c in 'A'..'Z') result = result * 26 + (c - 'A' + 1)
        return (result - 1).coerceAtLeast(0)
    }

    private fun parseCellRef(ref: String): Pair<Int, Int> {
        val col = ref.takeWhile { it.isLetter() }.uppercase()
        val row = ref.dropWhile { it.isLetter() }.toIntOrNull()?.minus(1) ?: 0
        return colLettersToIndex(col) to row.coerceAtLeast(0)
    }

    private fun applyCellFormat(wb: Workbook, sh: Sheet, range: String, fmt: JSONObject) {
        // Supports single cell or simple A1:B2 (best-effort single cell)
        val (col, row) = parseCellRef(range.substringBefore(":"))
        val r = sh.getRow(row) ?: sh.createRow(row)
        val c = r.getCell(col) ?: r.createCell(col)
        val style = wb.createCellStyle()
        if (fmt.has("bold") || fmt.has("italic") || fmt.has("font_size") || fmt.has("color")) {
            val font = wb.createFont()
            if (fmt.optBoolean("bold")) font.bold = true
            if (fmt.optBoolean("italic")) font.italic = true
            if (fmt.has("font_size")) font.fontHeightInPoints = fmt.optInt("font_size").toShort()
            val color = fmt.optString("color")
            if (color.startsWith("#")) font.color = IndexedColors.BLACK.index // keep simple
            style.setFont(font)
        }
        if (fmt.has("fill")) {
            style.fillForegroundColor = IndexedColors.GREY_25_PERCENT.index
            style.fillPattern = FillPatternType.SOLID_FOREGROUND
        }
        if (fmt.has("number_format")) style.dataFormat = wb.createDataFormat().getFormat(fmt.optString("number_format"))
        c.cellStyle = style
    }
}



object OfficeLibrary {
    @JvmStatic
    fun render(directory: java.io.File, arguments: org.json.JSONObject): String {
        return OfficeRenderer(directory, OfficeLimits()).render(arguments)
    }
}
