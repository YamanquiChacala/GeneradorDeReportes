import type { Icon } from "./types";

export * from "./color";
export * from "./constants";
export * from "./helpers";
export * from "./iconify";
export * from "./math";
export * from "./text";
export * from "./types";

export interface IconifyParams {
    iconName: Icon;
    color?: string;
    width?: number;
    height?: number;
    box?: boolean;
}
