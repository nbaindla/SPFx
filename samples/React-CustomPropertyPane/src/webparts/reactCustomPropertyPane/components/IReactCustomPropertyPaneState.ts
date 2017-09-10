import IPropsItem from './IPropsItem';

interface IReactCustomPropertyPaneState {
    isListPanel?: boolean;
    isLibraryPanel?: boolean;
    listDetails?: IPropsItem;
    libraryDetails?: IPropsItem;
}

export default IReactCustomPropertyPaneState;