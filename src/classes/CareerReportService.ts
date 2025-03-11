import * as XLSX from "xlsx";
import {WorkSheet} from "xlsx";
import {ValuedTableCoordinate} from "../types/ValuedTableCoordinate";
import {tableCoordinates} from "../variables/tableCoordinates";
import {promises as fsAsync} from "fs";
import {PDFDocument, rgb} from "pdf-lib";
import fontkit from "@pdf-lib/fontkit";
import {ImageElement} from "../types/ImageElement";
import {TextElement} from "../types/TextElement";
import generateChart from "../functions/generateChart";
import "dotenv/config"
import path from "node:path";

export default class CareerReportService {
    private worksheet!: WorkSheet;
    private valuedTableCoordinates!: ValuedTableCoordinate[]
    private pdfDoc!: PDFDocument
    private filePath!: string
    private studentName!: string

    public constructor() {
    }

    public async generateCareerReport(filePath: string, studentName: string) {
        this.filePath = filePath
        this.studentName = studentName
        this.initializeWorkSheet(filePath)
        this.initializeValuedTableCoordinates()
        await this.initializePDFDocument()
        await this.writeName()
        await this.writeTable()
        await this.writeChart()
        await this.writeProfession()
        await this.writeProfessionGroup()
        await this.writeBasicRequirements()
        await this.writeCareerPath()
        await this.writeClosing()
        await this.writeBatchNumber()
    }

    public async savePdf(outputPath: string) {
        await fsAsync.mkdir(path.dirname(outputPath), {recursive: true});
        await fsAsync.writeFile(outputPath, Buffer.from(await this.pdfDoc.save()));
    }

    private async writeBatchNumber() {
        await this.writePdf([{
            page: 2,
            params: [
                {
                    text: process.env.BATCH_NUMBER ? process.env.BATCH_NUMBER + " " : "1 ",
                    details: {
                        x: 117.62736525,
                        y: 612.8267791075,
                        size: 11,
                        color: rgb(1, 1, 1),
                        fontPath: "./public/fonts/arial/ARIALBD.ttf"
                    }
                }

            ]
        }])
    }

    private async writeClosing() {
        const pdf = await PDFDocument.load(await fsAsync.readFile(`./public/templates/06_closing.pdf`))
        const copiedPages = await this.pdfDoc.copyPages(pdf, pdf.getPageIndices());
        copiedPages.forEach((page) => this.pdfDoc.addPage(page));
    }

    private async writeCareerPath() {
        for (let i = 0; i < 3; i++) {
            const key = this.retrieveCellValue(6 + i, 1) + "-" + this.retrieveCellValue(6 + i, 2)
            const pdf = await PDFDocument.load(await fsAsync.readFile(`./public/templates/05_career path/${key}.pdf`))
            const copiedPages = await this.pdfDoc.copyPages(pdf, pdf.getPageIndices());
            copiedPages.forEach((page) => this.pdfDoc.addPage(page));
        }
    }

    private async writeBasicRequirements() {
        const pdf = await PDFDocument.load(await fsAsync.readFile(`./public/templates/04_basic requirements.pdf`))
        const copiedPages = await this.pdfDoc.copyPages(pdf, pdf.getPageIndices());
        copiedPages.forEach((page) => this.pdfDoc.addPage(page));
    }

    private async writeProfessionGroup() {
        const hashMap = new Map<string, boolean>()
        for (let i = 0; i < 3; i++) {
            const key = this.retrieveCellValue(6 + i, 1);
            if (!hashMap.has(key)) {
                const pdf = await PDFDocument.load(await fsAsync.readFile(`./public/templates/03_profesi/${key}.pdf`))
                const copiedPages = await this.pdfDoc.copyPages(pdf, pdf.getPageIndices());
                copiedPages.forEach((page) => this.pdfDoc.addPage(page));
                hashMap.set(key, true)
            }
        }
    }

    private async writeProfession() {
        const hashMap = new Map<string, boolean>()
        for (let i = 0; i < 3; i++) {
            const key = this.retrieveCellValue(6 + i, 2);
            if (!hashMap.has(key)) {
                const pdf = await PDFDocument.load(await fsAsync.readFile(`./public/templates/02_bidang/${key}.pdf`))
                const copiedPages = await this.pdfDoc.copyPages(pdf, pdf.getPageIndices());
                copiedPages.forEach((page) => this.pdfDoc.addPage(page));
                hashMap.set(key, true)
            }
        }
    }

    private async writeChart() {
        await this.writePdf(
            undefined,
            [
                {
                    page: 2,
                    params: [{
                        image: generateChart(this.filePath),
                        details: {
                            x: 72 + 12 * 3,
                            y: 328.56298934999995 - 12 * 8 - 700 * 0.18,
                            width: 1200 * 0.32,
                            height: 700 * 0.30
                        }
                    }]
                }
            ]
        )
    }

    private async writeTable() {
        await this.writePdf([
            ...this.valuedTableCoordinates.slice(0, 3).map((valuedCoordinate) => {
                return {
                    page: 2,
                    params: [
                        {
                            text: valuedCoordinate.C1.value,
                            details: {
                                x: valuedCoordinate.C1.x,
                                y: valuedCoordinate.C1.y,
                                size: 12,
                                color: rgb(1, 1, 1),
                                fontPath: "./public/fonts/montserrat/static/Montserrat-Bold.ttf"
                            }
                        },
                        {
                            text: valuedCoordinate.C2.value,
                            details: {
                                x: valuedCoordinate.C2.x,
                                y: valuedCoordinate.C2.y,
                                size: 12,
                                color: rgb(1, 1, 1),
                                fontPath: "./public/fonts/montserrat/static/Montserrat-Bold.ttf"
                            }
                        },
                        {
                            text: valuedCoordinate.C3.value.toString(),
                            details: {
                                x: valuedCoordinate.C3.x - 4.5 * 12,
                                y: valuedCoordinate.C3.y,
                                size: 12,
                                color: rgb(1, 1, 1),
                                fontPath: "./public/fonts/montserrat/static/Montserrat-Bold.ttf"
                            }
                        },
                    ]
                }
            }),
            ...this.valuedTableCoordinates.slice(3, this.valuedTableCoordinates.length).map((coordinate) => {
                return {
                    page: 2,
                    params: [
                        {
                            text: coordinate.C1.value,
                            details: {
                                x: coordinate.C1.x,
                                y: coordinate.C1.y,
                                size: 10,
                                color: rgb(1, 1, 1,),
                                fontPath: "./public/fonts/montserrat/static/Montserrat-Bold.ttf"
                            }
                        },
                        {
                            text: coordinate.C2.value,
                            details: {
                                x: coordinate.C2.x,
                                y: coordinate.C2.y,
                                size: 10,
                                color: rgb(1, 1, 1,),
                                fontPath: "./public/fonts/montserrat/static/Montserrat-Bold.ttf"
                            }
                        },
                        {
                            text: coordinate.C3.value.toString(),
                            details: {
                                x: coordinate.C3.x - 5 * 10,
                                y: coordinate.C3.y,
                                size: 10,
                                color: rgb(1, 1, 1,),
                                fontPath: "./public/fonts/montserrat/static/Montserrat-Bold.ttf"
                            }
                        },
                    ]
                }
            })
        ])
    }

    private async writeName() {
        await this.writePdf([
            {
                page: 2,
                params: [
                    {
                        text: 'Halo,',
                        details: {
                            x: 57.75,
                            y: 650.5407235,
                            size: 25,
                            color: rgb(1, 1, 1),
                            fontPath: './public/fonts/montserrat/static/Montserrat-Regular.ttf'
                        }
                    },
                    {
                        text: `${this.studentName.split(" ").length > 2 ?
                            this.studentName.split(" ").splice(0, 2).join(" ") : this.studentName}!`,
                        details: {
                            x: 130.97972945200002,
                            y: 650.5407235,
                            size: 25,
                            color: rgb(1, 1, 1),
                            fontPath: "./public/fonts/montserrat/static/Montserrat-Bold.ttf"
                        }
                    }
                ]
            }
        ]);
    }

    private async initializePDFDocument() {
        this.pdfDoc = await PDFDocument.load(await fsAsync.readFile("./public/template.pdf"));
        this.pdfDoc.registerFontkit(fontkit)
    }

    private async writePdf(textElements?: TextElement[], imageElements?: ImageElement[]): Promise<void> {
        const pages = this.pdfDoc.getPages();
        if (textElements) {
            for (const textElement of textElements) {
                const page = pages[textElement.page - 1];
                for (const params of textElement.params) {
                    const {fontPath, ...restDetails} = params.details;

                    page.drawText(params.text, {
                        ...restDetails, ...fontPath ? {
                            font: await this.pdfDoc.embedFont(await fsAsync.readFile(fontPath))
                        } : {}
                    });
                }
            }
        }
        if (imageElements) {
            for (const imageElement of imageElements) {
                const page = pages[imageElement.page - 1];
                for (const param of imageElement.params) {
                    page.drawImage(await this.pdfDoc.embedPng(param.image), param.details);
                }
            }
        }
    }

    private initializeValuedTableCoordinates() {
        this.valuedTableCoordinates = []
        for (let i = 0; i < tableCoordinates.length; i++) {
            this.valuedTableCoordinates.push(
                {
                    C1: {...tableCoordinates[i].C1, value: this.retrieveCellValue(6 + i, 1)},
                    C2: {...tableCoordinates[i].C2, value: this.retrieveCellValue(6 + i, 2)},
                    C3: {
                        ...tableCoordinates[i].C3,
                        value: Number(Number(this.retrieveCellValue(6 + i, 3)).toFixed(6))
                    }
                }
            )
        }
    }

    private initializeWorkSheet(filePath: string) {
        const workbook = XLSX.readFile(filePath);
        this.worksheet = workbook.Sheets[workbook.SheetNames[0]];
    }

    private retrieveCellValue(row: number, column: number): string {
        const cell = this.worksheet[XLSX.utils.encode_cell({r: row - 1, c: column - 1})];
        if (cell == null) {
            throw new Error("Invalid row and column!")
        }
        return cell.v
    }
}