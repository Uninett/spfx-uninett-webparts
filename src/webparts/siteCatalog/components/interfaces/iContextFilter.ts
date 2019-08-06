export interface IContextFilter {
    fieldName: string; // column.key
    filterValue?: any; // Bibsys
    filterText?: string; // Bibsys
    isFiltered?: boolean; // true
    direction?: string;
}