import * as React from "react";
import styles from "./Hello.module.scss";
import { IHelloProps } from "./IHelloProps";
import { escape } from "@microsoft/sp-lodash-subset";
import SPService from "../../../_services/SPServices";

const testFields = [
  {
    Title: "TextField",
    FieldTypeKind: 2,
  },
  {
    Title: "Number",
    FieldTypeKind: 3,
  },
  {
    Title: "Date",
    FieldTypeKind: 4,
  },
  {
    Title: "User",
    FieldTypeKind: 20,
  },
];

export default class Hello extends React.Component<IHelloProps, {}> {
  public spService = new SPService(this.props.context);

  componentDidMount(): void {
    // this.spService.createList("SampleTestList");
    // this.spService.createSiteField("fieldone","SampleTestList")
    // this.spService.createSiteForAList("Column_one","SampleTestList")

    this.spService.createFieldsForAList("SampleTestList", testFields);
  }
  public render(): React.ReactElement<IHelloProps> {
    return (
      <div className={styles.hello}>
        <div className={styles.container}>
          <div className={styles.row}>
            <div className={styles.column}>
              <span className={styles.title}>Welcome to SharePoint!</span>
              <p className={styles.subTitle}>
                Customize SharePoint experiences using Web Parts.
              </p>
              <p className={styles.description}>
                {escape(this.props.description)}
              </p>
              <a href="https://aka.ms/spfx" className={styles.button}>
                <span className={styles.label}>Learn more</span>
              </a>
            </div>
          </div>
        </div>
      </div>
    );
  }
}
