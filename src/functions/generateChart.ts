import * as XLSX from "xlsx";
import {CanvasRenderingContext2D, createCanvas, registerFont} from "canvas";
import * as d3 from "d3";

registerFont("./public/fonts/montserrat/static/Montserrat-Bold.ttf", {family: "Montserrat"});

const width: number = 1200;
const height: number = 700;
const marginLeft = 450;
const mainColor = "red";
const topColor = "rgb(255, 120, 120)";
const sideColor = "rgb(180, 0, 0)";
const depth = 10;

const canvas = createCanvas(width, height);
const context: CanvasRenderingContext2D = canvas.getContext("2d");

export default function generateChart(filePath: string) {
    const workbook = XLSX.readFile(filePath);
    const worksheet = workbook.Sheets[workbook.SheetNames[0]];

    function getCellValue(row: number, col: number) {
        const cellAddress = XLSX.utils.encode_cell({r: row - 1, c: col - 1});
        const cell = worksheet[cellAddress];
        if (cell == null) {
            throw new Error("Invalid row and column!")
        }
        return cell.v
    }

    const fields: string[] = []
    const scores: number[] = []
    for (let i = 0; i < 10; i++) {
        fields.push(getCellValue(6 + i, 2) + " " + getCellValue(6 + i, 1).slice(0, -1))
        scores.push(Number(getCellValue(6 + i, 3).toFixed(6)))
    }

    const x = d3.scaleLinear().domain([0, 1]).range([marginLeft, 1100]);
    const y = d3.scaleBand<string>().domain(fields).range([100, 650]).padding(0.5);

    context.fillStyle = "white";
    context.fillRect(0, 0, width, height);

    context.fillStyle = "black";
    context.font = "bold 24px";
    context.textAlign = "center";
    context.fillText("10 HASIL KARIR TERTINGGI", width / 2, 40);

    context.fillStyle = "gray";
    context.font = "bold 14px";
    context.textAlign = "center";

    for (let value = 0; value <= 0.8; value += 0.2) {
        const xPos = x(value);

        context.strokeStyle = "lightgray";
        context.beginPath();
        context.moveTo(xPos, 80);
        context.lineTo(xPos, 650);
        context.stroke();

        context.fillStyle = "black";
        context.fillText(value.toFixed(1), xPos, 650);
    }

    scores.forEach((d, i) => {
        const barY = y(fields[i]);
        const barHeight = y.bandwidth();
        const barWidth = x(d) - marginLeft;

        if (barY !== undefined && barHeight !== undefined) {
            context.fillStyle = "black";
            context.font = "bold 16px";
            context.textAlign = "right";
            context.fillText(fields[i], marginLeft - 10, barY + barHeight / 2 + 5);

            context.fillStyle = mainColor;
            context.fillRect(marginLeft, barY, barWidth, barHeight);

            context.fillStyle = topColor;
            context.beginPath();
            context.moveTo(marginLeft, barY);
            context.lineTo(marginLeft + depth, barY - depth);
            context.lineTo(marginLeft + barWidth + depth, barY - depth);
            context.lineTo(marginLeft + barWidth, barY);
            context.closePath();
            context.fill();

            context.fillStyle = sideColor;
            context.beginPath();
            context.moveTo(marginLeft + barWidth, barY);
            context.lineTo(marginLeft + barWidth + depth, barY - depth);
            context.lineTo(marginLeft + barWidth + depth, barY + barHeight - depth);
            context.lineTo(marginLeft + barWidth, barY + barHeight);
            context.closePath();
            context.fill();

            context.fillStyle = "black";
            context.font = "bold 16px";
            context.textAlign = "left";
            context.fillText(d.toString(), marginLeft + barWidth + 15, barY + barHeight / 2 + 5);
        }
    });

    return canvas.toBuffer("image/png");
}