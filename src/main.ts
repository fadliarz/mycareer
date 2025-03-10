import * as fs from "node:fs";
import CareerReportService from "./classes/CareerReportService";

async function main() {
    const filePaths: string[] = fs.readdirSync("./input").filter((file) => file.endsWith(".xlsx"))
    let count: number = 0;
    let total = filePaths.length;
    const globalStartTimeInSec: number = new Date().getTime() / 1000;
    for (const filePath of filePaths) {
        const startTimeInSec: number = new Date().getTime() / 1000;
        const studentName = filePath.split("_")[1];
        const careerReportService = new CareerReportService()
        await careerReportService.generateCareerReport(`./input/${filePath}`, studentName)
        await careerReportService.savePdf(`./output/${studentName}.pdf`)

        console.log(`Processed file ${filePath} in ${(new Date().getTime() / 1000 - startTimeInSec).toFixed(2)} s * (${(++count / total * 100).toFixed(2)}%)`)
    }
    console.log(`Successfully generated PDFs from ${total} files in ${(new Date().getTime() / 1000 - globalStartTimeInSec).toFixed(2)} s`)
}

main().catch(
    (error) => {
        console.log("Failed to modify PDF! Error: ", error)
    }
)