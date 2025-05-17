import { DataViewObjectPropertyReference } from "./common";

export class SubtotalProperties {
  public static readonly ObjectSubTotals: string = "subTotals";

  public static rowSubtotals: DataViewObjectPropertyReference<boolean> = {
    propertyIdentifier: {
      objectName: SubtotalProperties.ObjectSubTotals,
      propertyName: "rowSubtotals",
    },
    defaultValue: true,
  };

  public static rowSubtotalsPerLevel: DataViewObjectPropertyReference<boolean> =
    {
      propertyIdentifier: {
        objectName: SubtotalProperties.ObjectSubTotals,
        propertyName: "perRowLevel",
      },
      defaultValue: false,
    };

  public static columnSubtotals: DataViewObjectPropertyReference<boolean> = {
    propertyIdentifier: {
      objectName: SubtotalProperties.ObjectSubTotals,
      propertyName: "columnSubtotals",
    },
    defaultValue: true,
  };

  public static columnSubtotalsPerLevel: DataViewObjectPropertyReference<boolean> =
    {
      propertyIdentifier: {
        objectName: SubtotalProperties.ObjectSubTotals,
        propertyName: "perColumnLevel",
      },
      defaultValue: false,
    };

  public static levelSubtotalEnabled: DataViewObjectPropertyReference<boolean> =
    {
      propertyIdentifier: {
        objectName: SubtotalProperties.ObjectSubTotals,
        propertyName: "levelSubtotalEnabled",
      },
      defaultValue: true,
    };
}
