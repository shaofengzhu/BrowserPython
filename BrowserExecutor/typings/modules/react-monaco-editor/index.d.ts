
declare module 'react-monaco-editor' {
    import * as React from 'react';

    interface MonacoEditorProps{
        width: number | string;
        height: number | string;
        value: string;
        defaultValue: string;
        language: string;
        editorDidMount: (e:any) => void;
        editorWillMount: (e:any) => void;
        onChange: (e:string) => void;
    }

    class MonacoEditor extends React.Component<any, any>{    
    }

    export default MonacoEditor;
}
