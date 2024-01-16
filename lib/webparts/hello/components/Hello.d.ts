/// <reference types="react" />
import * as React from "react";
import { IHelloProps } from "./IHelloProps";
import SPService from "../../../_services/SPServices";
export default class Hello extends React.Component<IHelloProps, {}> {
    spService: SPService;
    componentDidMount(): void;
    render(): React.ReactElement<IHelloProps>;
}
