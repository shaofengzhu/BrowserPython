export class Utility {
    static log(message: string): void {
        console.log(message);
        var div = document.createElement("div");
        div.innerText = message;
        document.getElementById("DivLog").appendChild(div);
    }

    static isNullOrUndefined(obj: any): boolean {
        return typeof (obj) === "undefined" || obj === null;
    }

    static isNullOrEmptyString(str: string): boolean {
        if (Utility.isNullOrUndefined(str))
            return true;
        return str.length == 0;
    }
}
