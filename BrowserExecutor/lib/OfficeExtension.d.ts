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
}
