import React, { useCallback, useState } from "react";
import { read, utils } from "xlsx";
import { saveAs } from "file-saver";
import {
    Document,
    Packer,
    Paragraph,
    TextRun,
    ShadingType,
} from "docx";

import "./Excel.css"


const ExcelDropToWord = () => {
    const [status, setStatus] = useState("Drag and drop Excel file here");

    const handleDrop = useCallback(async (e) => {
        e.preventDefault();
        const file = e.dataTransfer.files[0];

        if (!file || !file.name.endsWith(".xlsx")) {
            setStatus("❌ Please upload a valid .xlsx file.");
            return;
        }

        setStatus("📄 Processing Excel...");

        try {
            const buffer = await file.arrayBuffer();
            const workbook = read(buffer);
            const sheet = workbook.Sheets["Open Items List"];
            if (!sheet) {
                setStatus("❌ Sheet 'Open Items List' not found.");
                return;
            }

            const data = utils.sheet_to_json(sheet, { header: 1 }).slice(2);
            const allTexts = data.map((row) => row[6]).filter(Boolean);

            const exceptions = allTexts.filter((t) =>
                t.toLowerCase().includes("exception")
            );
            const clarifications = allTexts.filter((t) =>
                t.toLowerCase().includes("clarification")
            );

            const makeBullet = (text) => {
                const parts = text.toString().split(/(Exception:|Clarification:)/i);
                const header = parts[0]?.trim();
                const type = (parts[1] || "").replace(":", "").trim();  // 🧼 cleaned
                const message = parts[2]?.trim() || "";

                return new Paragraph({
                    bullet: { level: 0 },
                    children: [
                        new TextRun({
                            text: `${header}`,
                            font: "arial narrow",
                            size: 20,
                            shading: {
                                type: ShadingType.CLEAR,
                                color: "auto",
                            },
                        }),
                        new TextRun({ break: 1 }), // ⬅️ forces a new line
                        new TextRun({
                            text: `${type}: ${message}`,
                            font: "arial narrow",
                            size: 20,
                            shading: {
                                type: ShadingType.CLEAR,
                                color: "auto",
                            },
                        }),
                    ],
                });
            };



            const doc = new Document({
                sections: [
                    {
                        children: [
                            new Paragraph({
                                indent: { left: 720 },
                                children: [
                                    new TextRun({
                                        text: "Exception and Clarification",
                                        font: "arial narrow",
                                        shading: {
                                            type: ShadingType.CLEAR,
                                            color: "auto",
                                            fill: "FFFF00", // Yellow background
                                        },
                                        bold: true,
                                        size: 20,
                                    }),
                                ],
                            }),
                            new Paragraph({
                                indent: { left: 720 },
                                children: [
                                    new TextRun({
                                        text: "Exceptions:", font: "Arial Narrow", shading: {
                                            type: ShadingType.CLEAR,
                                            color: "auto",
                                            fill: "FFFF00", // Yellow background
                                        }, size: 20
                                    }),
                                ],
                            }),
                            ...exceptions.map(makeBullet),
                            new Paragraph({
                                indent: { left: 720 },
                                children: [
                                    new TextRun({
                                        font: "arial narrow",
                                        text: "Clarifications:", shading: {
                                            type: ShadingType.CLEAR,
                                            color: "auto",
                                            fill: "FFFF00", // Yellow background
                                        }, size: 20
                                    }),
                                ],
                            }),
                            ...clarifications.map(makeBullet),
                        ],
                    },
                ],
            });

            const blob = await Packer.toBlob(doc);
            saveAs(blob, "Exceptions_and_Clarifications.docx");

            setStatus("✅ Word document created!");


            setTimeout(() => {
                setStatus("Drag and drop Excel file here");
            }, 2000);
        } catch (err) {
            console.error(err);
            setStatus("❌ Error processing Excel.");
        }
    }, []);

    const handleDragOver = (e) => e.preventDefault();

    return (

        <div className="container">
            <h1 className="title">📘 Excel to Word Converter</h1>
            <p className="subtitle">
                This tool converts structured data from an Excel file into a formatted Word document, making it easy to generate reports.
            </p>

            <div className="drop-area" onDrop={handleDrop} onDragOver={handleDragOver}>
                <h2>📥 Drag & Drop Excel (.xlsx) File</h2>
                <p className="status">{status}</p>
            </div>

            <footer className="footer">Made by Anirban Roy</footer>
        </div>


    );
};

export default ExcelDropToWord;
