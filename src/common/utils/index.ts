import type { Icon } from "../constants";

export * from "./color";
export * from "./helper";
export * from "./iconify";
export * from "./math";
export * from "./text";

export interface IconifyParams {
    iconName: Icon;
    color?: string;
    width?: number;
    height?: number;
    box?: boolean;
}
