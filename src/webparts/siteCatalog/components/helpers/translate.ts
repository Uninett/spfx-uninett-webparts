import * as strings from 'SiteCatalogWebPartStrings';

let T = (colName: string):string => {
    let translation = strings[colName];

    if (!translation)
        return colName;

    return translation;
};

export { T };