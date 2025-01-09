import { IRequest } from "../ds";

export interface ITab {
    getProps(): { [key: string]: string | number | boolean }
    getRequests(): IRequest[];
}