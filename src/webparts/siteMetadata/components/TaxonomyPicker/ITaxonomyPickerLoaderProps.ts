import { ITaxonomyPickerProps } from "react-taxonomypicker";
import { WebPartContext } from "@microsoft/sp-webpart-base";

export interface ITaxonomyPickerLoaderProps extends ITaxonomyPickerProps {
    context: WebPartContext;
}