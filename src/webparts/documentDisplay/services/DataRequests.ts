import { WebPartContext } from "@microsoft/sp-webpart-base";

export const getFileTypeUrl = (fileName: string) => {
    const fileExt = fileName.substring(fileName.indexOf('.'));
    let iconName = 'sharepoint';
    let fileUrl;
    let isOfficeFile = true;

    switch(fileExt){
        case '.doc':
        case '.docx':
        case '.rtf':
        case '.dotx':
            iconName = 'word'
        break;
        case '.ppt':
        case '.pptx':
        case '.potx':
        case '.ppsx':
            iconName = 'powerpoint'
        break;
        case '.xls':
        case '.xlsx':
        case '.xltx':
        case '.csv':
            iconName = 'excel';
        break;
        case '.vsd':
        case '.vsdx':
        case '.vssx':
        case '.vstx':
            iconName = 'visio';
        break;
        case '.mpp':
        case '.mppx':
        case '.mpt':
            iconName = 'project';
        break;
        case '.onetoc':
        case '.one':
            iconName = 'onenote';
        break;
        case '.pdf':
            isOfficeFile = false;
            fileUrl = require('../assets/pdf.png');
        break;
        case '.gif':
        case '.jpg':
        case '.jpeg':
        case '.bmp':
        case '.dib':
        case '.tif':
        case '.tiff':
        case '.ico':
        case '.png':
        case '.jxr':
        case '.svg':
            isOfficeFile = false;
            fileUrl = require('../assets/pic-icon.png');
        break;
        case '.aspx':
            if (fileName.toLowerCase().indexOf('onenote') !== -1) iconName = 'onenote';
            else iconName = 'sharepoint';
        break;
        default:
            iconName = 'sharepoint';
        break;
    }

    if (isOfficeFile) fileUrl = `https://res-1.cdn.office.net/files/fabric-cdn-prod_20230815.002/assets/brand-icons/product/svg/${iconName}_48x1.svg`;

    return fileUrl;
};

export const getGraphMemberOf = async (context: WebPartContext) : Promise <string> =>{
    const graphResponse = await context.msGraphClientFactory.getClient('3');
    const graphUrl = '/me/transitiveMemberOf/microsoft.graph.group';
    const memberOfGraph = await graphResponse
        .api(graphUrl)
        .header('ConsistencyLevel', 'eventual')
        .count(true)
        .select('displayName')
        .top(500)
        .get();

    return memberOfGraph;
};

export const isFromTargetAudience = (context: WebPartContext, graphResponse: any, wpTargetAudience: any) => {
    
    for (const audience of wpTargetAudience){
        if (context.pageContext.user.email === audience.email)
            return true;
    }
    
    const userGroups = [];
    for (const group of graphResponse){
        userGroups[group.displayName] = group.displayName;
    }
    for (const audience of wpTargetAudience){
        if (userGroups[audience.fullName])
            return true;
    }
    return false;
};