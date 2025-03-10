import {getDocument, GlobalWorkerOptions} from 'pdfjs-dist';

GlobalWorkerOptions.workerSrc = require.resolve('pdfjs-dist/legacy/build/pdf.worker.js');

export default async function getTextCoordinates(pdfPath: string) {
    const loadingTask = getDocument(pdfPath);
    const pdf = await loadingTask.promise;

    const page = await pdf.getPage(2);
    const textContent = await page.getTextContent();

    textContent.items.forEach((item) => {
        if ('str' in item) {
            console.log('Text:', item.str);
            console.log('Coordinates (x,y):', item.transform[4], item.transform[5]);
            console.log('-------------------');
        }
    });
}