package com.teamshryne.kinooffice

import org.apache.poi.xslf.usermodel.XMLSlideShow
import org.apache.poi.xssf.usermodel.XSSFWorkbook
import org.apache.poi.xwpf.usermodel.XWPFDocument
import org.json.JSONArray
import org.json.JSONObject
import org.junit.Assert.assertTrue
import org.junit.Test
import java.io.File

class OfficeRendererTest {
    @Test
    fun createsWordDocument() {
        val dir = File(System.getProperty("java.io.tmpdir"), "office-test-word").apply { mkdirs() }
        val args = JSONObject()
            .put("format", "docx")
            .put("operation", "create")
            .put("title", "Sample Word")
            .put("markdown", "# Title\n\nHello from Kino.\n\n- One\n- Two\n\n| A | B |\n|---|---|\n| 1 | 2 |")
        val filename = OfficeRenderer(dir, OfficeLimits()).render(args)
        val file = File(dir, filename)
        assertTrue(file.exists() && file.length() > 0)
        XWPFDocument(file.inputStream()).use { doc ->
            assertTrue(doc.paragraphs.isNotEmpty())
            assertTrue(doc.tables.isNotEmpty())
        }
    }

    @Test
    fun createsSpreadsheet() {
        val dir = File(System.getProperty("java.io.tmpdir"), "office-test-xlsx").apply { mkdirs() }
        val rows = JSONArray()
            .put(JSONArray().put("Month").put("Revenue"))
            .put(JSONArray().put("Jan").put(1200))
            .put(JSONArray().put("Feb").put(1500))
        val sheets = JSONArray().put(JSONObject().put("name", "Sales").put("rows", rows))
        val args = JSONObject()
            .put("format", "xlsx")
            .put("operation", "create")
            .put("title", "Sample Excel")
            .put("sheets", sheets)
        val filename = OfficeRenderer(dir, OfficeLimits()).render(args)
        val file = File(dir, filename)
        assertTrue(file.exists() && file.length() > 0)
        XSSFWorkbook(file.inputStream()).use { wb ->
            assertTrue(wb.numberOfSheets > 0)
            assertTrue(wb.getSheetAt(0).physicalNumberOfRows > 0)
        }
    }

    @Test
    fun createsPresentation() {
        val dir = File(System.getProperty("java.io.tmpdir"), "office-test-pptx").apply { mkdirs() }
        val slides = JSONArray()
            .put(JSONObject().put("title", "Slide 1").put("bullets", JSONArray().put("Bullet 1").put("Bullet 2")))
        val args = JSONObject()
            .put("format", "pptx")
            .put("operation", "create")
            .put("title", "Sample PPT")
            .put("slides", slides)
        val filename = OfficeRenderer(dir, OfficeLimits()).render(args)
        val file = File(dir, filename)
        assertTrue(file.exists() && file.length() > 0)
        XMLSlideShow(file.inputStream()).use { ppt ->
            assertTrue(ppt.slides.isNotEmpty())
        }
    }
}
