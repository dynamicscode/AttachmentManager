import { mergeStyleSets } from "office-ui-fabric-react/lib/Styling";

export const classNames = mergeStyleSets({
    fullWidthControl: {
        width: '100%'
    },
    fileIcon: {
        fontSize: 20
    },
    wrapper: {
        height: '60vh',
        position: 'relative'
    },
    filter: {
        paddingBottom: 20,
        maxWidth: 300
    },
    header: {
        margin: 0
    },
    row: {
        display: 'inline-block'
    }
});