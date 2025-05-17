import powerbi from "powerbi-visuals-api";
import DataViewObjectPropertyIdentifier = powerbi.DataViewObjectPropertyIdentifier;

/** Defines a selector for content, including data-, metadata, and user-defined repetition. */
export interface Selector {
  /** Metadata-bound repetition selection.  Refers to a DataViewMetadataColumn queryName. */
  metadata?: string;

  /** User-defined repetition selection. */
  id?: string;
}

export interface DataViewObjectPropertyReference<T> {
  /** Property identifier that holds the Value. Only static properties (Null Selector) are supported */
  propertyIdentifier?: DataViewObjectPropertyIdentifier;

  /** Value to use if the PropertyDefinition does not exist */
  defaultValue: T;
}
