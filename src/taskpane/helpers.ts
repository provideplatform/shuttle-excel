import { Model } from "@provide/types";

export function snakeToCamel<T>(from: any): T {
    const model = new Model();
    model.unmarshal(JSON.stringify(from));
    return model as any as T;
}
