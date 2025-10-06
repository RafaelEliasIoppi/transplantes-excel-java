package com.seuprojeto.transplantes.controller;

import java.io.IOException;
import java.util.ArrayList;
import java.util.List;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.http.HttpStatus;
import org.springframework.http.MediaType;
import org.springframework.http.ResponseEntity;
import org.springframework.web.bind.annotation.PostMapping;
import org.springframework.web.bind.annotation.RequestBody;
import org.springframework.web.bind.annotation.RequestParam;
import org.springframework.web.bind.annotation.RestController;
import org.springframework.web.multipart.MultipartFile;

import com.itextpdf.text.DocumentException;
import com.seuprojeto.transplantes.service.RelatorioPdfService;

@RestController
public class ExcelController {

    @Autowired
    private RelatorioPdfService relatorioPdfService;

    @PostMapping("/relatorio-pdf")
    public ResponseEntity<byte[]> gerarPdf(@RequestBody List<String> dados) {
        try {
            byte[] pdf = relatorioPdfService.gerarRelatorio(dados);
            return ResponseEntity.ok()
                    .header("Content-Disposition", "attachment; filename=relatorio.pdf")
                    .contentType(MediaType.APPLICATION_PDF)
                    .body(pdf);
        } catch (DocumentException e) {
            return ResponseEntity.status(HttpStatus.INTERNAL_SERVER_ERROR).body(null);
        }
    }

    @PostMapping("/update")
    public ResponseEntity<List<String>> extraiUltimosValoresPorColunaDeTodasAsAbas(@RequestParam("file") MultipartFile file) {
        List<String> relatorio = new ArrayList<>();

        try (Workbook workbook = new XSSFWorkbook(file.getInputStream())) {
            FormulaEvaluator evaluator = workbook.getCreationHelper().createFormulaEvaluator();
            int totalAbas = workbook.getNumberOfSheets();

            for (int abaIndex = 0; abaIndex < totalAbas; abaIndex++) {
                Sheet sheet = workbook.getSheetAt(abaIndex);
                String nomeAba = sheet.getSheetName();
                Row cabecalho = sheet.getRow(0);

                if (cabecalho == null) {
                    relatorio.add(nomeAba + " - (sem cabeçalho)");
                    continue;
                }

                int totalColunas = cabecalho.getLastCellNum();

                for (int col = 0; col < totalColunas; col++) {
                    String titulo = getValorCelula(cabecalho.getCell(col), evaluator);
                    String ultimoValor = "(não encontrado)";

                    for (int rowIndex = sheet.getLastRowNum(); rowIndex > 0; rowIndex--) {
                        Row row = sheet.getRow(rowIndex);
                        if (row != null) {
                            Cell cell = row.getCell(col);
                            if (cell != null && cell.getCellType() != CellType.BLANK) {
                                ultimoValor = getValorCelula(cell, evaluator);
                                break;
                            }
                        }
                    }

                    relatorio.add(nomeAba + " - " + titulo + ": " + ultimoValor);
                }
            }

            return ResponseEntity.ok(relatorio);

        } catch (IOException e) {
            return ResponseEntity.badRequest().body(List.of("Erro ao processar o arquivo: " + e.getMessage()));
        }
    }

    private String getValorCelula(Cell cell, FormulaEvaluator evaluator) {
        if (cell == null) return "(vazio)";

        CellType type = cell.getCellType();
        if (type == CellType.FORMULA) {
            CellValue evaluated = evaluator.evaluate(cell);
            if (evaluated == null) return "(erro ao avaliar fórmula)";
            type = evaluated.getCellType();

            return switch (type) {
                case STRING -> evaluated.getStringValue();
                case NUMERIC -> DateUtil.isCellDateFormatted(cell)
                        ? cell.getDateCellValue().toString()
                        : String.valueOf(evaluated.getNumberValue());
                case BOOLEAN -> String.valueOf(evaluated.getBooleanValue());
                case ERROR -> "(erro)";
                default -> "(tipo não reconhecido)";
            };
        } else {
            return switch (type) {
                case STRING -> cell.getStringCellValue();
                case NUMERIC -> DateUtil.isCellDateFormatted(cell)
                        ? cell.getDateCellValue().toString()
                        : String.valueOf(cell.getNumericCellValue());
                case BOOLEAN -> String.valueOf(cell.getBooleanCellValue());
                case BLANK -> "(em branco)";
                case ERROR -> "(erro)";
                default -> "(tipo não reconhecido)";
            };
        }
    }
}
