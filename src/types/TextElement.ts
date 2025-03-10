import {RGB} from "pdf-lib";

export type TextElement = {
    page: number,
    params: {
        text: string,
        details: {
            x: number,
            y: number,
            size: number,
            color: RGB,
            fontPath?: string
        }
    }[]
};
