package com.seuprojeto.transplantes.service;

import com.itextpdf.text.*;
import com.itextpdf.text.pdf.PdfWriter;
import org.springframework.stereotype.Service;

import java.io.ByteArrayOutputStream;
import java.util.List;

@Service
public class RelatorioPdfService {

    public byte[] gerarRelatorio(List<String> dados) throws DocumentException {
        Document document = new Document();
        ByteArrayOutputStream outputStream = new ByteArrayOutputStream();

        PdfWriter.getInstance(document, outputStream);
        document.open();

        // Título
        Font tituloFont = FontFactory.getFont(FontFactory.HELVETICA_BOLD, 16);
        Paragraph titulo = new Paragraph("Relatório de Últimos Valores por Coluna", tituloFont);
        titulo.setAlignment(Element.ALIGN_CENTER);
        titulo.setSpacingAfter(20);
        document.add(titulo);

        // Data
        Paragraph data = new Paragraph("Gerado em: " + java.time.LocalDateTime.now());
        data.setAlignment(Element.ALIGN_RIGHT);
        data.setSpacingAfter(10);
        document.add(data);

        // Conteúdo
        Font conteudoFont = FontFactory.getFont(FontFactory.HELVETICA, 12);
        for (String linha : dados) {
            Paragraph p = new Paragraph(linha, conteudoFont);
            p.setSpacingAfter(5);
            document.add(p);
        }

        document.close();
        return outputStream.toByteArray();
    }
}