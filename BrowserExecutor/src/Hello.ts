import * as React from 'react'

export class Hello extends React.Component<any, any>{
    constructor(props: any){
        super(props);
    }

    public render(){
        return React.createElement('div', null, "Hello, world");
    }
}

