export declare namespace OfficeExtension {
    interface IPromise<R> {
        then<U>(onFulfilled?: (value: R) => IPromise<U>, onRejected?: (error: any) => IPromise<U>): IPromise<U>;
        then<U>(onFulfilled?: (value: R) => IPromise<U>, onRejected?: (error: any) => U): IPromise<U>;
        then<U>(onFulfilled?: (value: R) => IPromise<U>, onRejected?: (error: any) => void): IPromise<U>;
        then<U>(onFulfilled?: (value: R) => U, onRejected?: (error: any) => IPromise<U>): IPromise<U>;
        then<U>(onFulfilled?: (value: R) => U, onRejected?: (error: any) => U): IPromise<U>;
        then<U>(onFulfilled?: (value: R) => U, onRejected?: (error: any) => void): IPromise<U>;
        catch<U>(onRejected?: (error: any) => IPromise<U>): IPromise<U>;
        catch<U>(onRejected?: (error: any) => U): IPromise<U>;
        catch<U>(onRejected?: (error: any) => void): IPromise<U>;
    }
    class Promise<R> implements IPromise<R> {
        constructor(func: (resolve, reject) => void);
        static all<U>(promises: OfficeExtension.IPromise<U>[]): IPromise<U[]>;
        static resolve<U>(value: U): IPromise<U>;
        static reject<U>(error: any): IPromise<U>;
        then<U>(onFulfilled?: (value: R) => IPromise<U>, onRejected?: (error: any) => IPromise<U>): IPromise<U>;
        then<U>(onFulfilled?: (value: R) => IPromise<U>, onRejected?: (error: any) => U): IPromise<U>;
        then<U>(onFulfilled?: (value: R) => IPromise<U>, onRejected?: (error: any) => void): IPromise<U>;
        then<U>(onFulfilled?: (value: R) => U, onRejected?: (error: any) => IPromise<U>): IPromise<U>;
        then<U>(onFulfilled?: (value: R) => U, onRejected?: (error: any) => U): IPromise<U>;
        then<U>(onFulfilled?: (value: R) => U, onRejected?: (error: any) => void): IPromise<U>;
        catch<U>(onRejected?: (error: any) => IPromise<U>): IPromise<U>;
        catch<U>(onRejected?: (error: any) => U): IPromise<U>;
        catch<U>(onRejected?: (error: any) => void): IPromise<U>;
    }
    class RichApiMessageUtility {
        private static OfficeJsErrorCode_ooeNoCapability;
        private static OfficeJsErrorCode_ooeActivityLimitReached;
        static buildResponseOnSuccess(responseBody: string, responseHeaders: {
            [name: string]: string;
        }): any;
        static buildResponseOnError(errorCode: number, message: string): any;
        static buildRequestMessageSafeArray(customData: string, requestFlags: number, method: string, path: string, headers: {
            [name: string]: string;
        }, body: string): Array<any>;
        static getResponseBody(result: any): string;
        static getResponseHeaders(result: any): {
            [name: string]: string;
        };
        static getResponseBodyFromSafeArray(data: Array<any>): string;
        static getResponseHeadersFromSafeArray(data: Array<any>): {
            [name: string]: string;
        };
        static getResponseStatusCode(result: any): number;
        static getResponseStatusCodeFromSafeArray(data: Array<any>): number;
    }    
}
