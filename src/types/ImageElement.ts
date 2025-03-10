export type ImageElement = {
    page: number,
    params: {
        image: Buffer,
        details: {
            x: number,
            y: number,
            width: number,
            height: number
        }
    }[]
};
