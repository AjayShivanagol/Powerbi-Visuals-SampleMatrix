"use strict";
import powerbi from "powerbi-visuals-api";
import md5 from "md5";

type CommentObject = {
  id: string;
  column: string;
  comment: string;
  filterId: string;
  createdAt?: string;
  modifiedAt?: string;
  color?: string;
};

export class MatrixDataviewHtmlFormatter {
  public static config = {
    redBarThreshold: 0,
    redBarColumnsIndex: [0, 2],
    commentApiUrl: "https://default-url/api/comment",
  };

  public static setConfig(config: {
    // ✅ CONFIG: Set threshold and column indexes where red bar should appear
    redBarThreshold: number;
    redBarColumnsIndex: number[];
    commentApiUrl: string;
  }) {
    this.config = config;
  }

  private static rowIndex = 0;
  private static commentMap: Map<string, CommentObject[]> = new Map();

  static setCommentMap(comments: CommentObject[]) {
    const grouped = new Map<string, CommentObject[]>();
    for (const comment of comments) {
      if (!grouped.has(comment.filterId)) {
        grouped.set(comment.filterId, []);
      }
      grouped.get(comment.filterId).push(comment);
    }
    this.commentMap = grouped;
  }

  public static formatDataViewMatrix(
    matrix: powerbi.DataViewMatrix
  ): HTMLElement {
    const htmlElement = document.createElement("div");
    htmlElement.classList.add("datagrid");

    const tableElement = document.createElement("table");
    tableElement.className = "matrixTable";

    const tbodyElement = document.createElement("tbody");
    const levelToColumnNodesMap: any[][] = [];

    const hasRealColumnHierarchy =
      matrix.columns?.root?.children &&
      matrix.columns.root.children.length > 0 &&
      matrix.columns.root.children[0].value != null;

    if (hasRealColumnHierarchy) {
      this.countColumnNodeLeaves(matrix.columns.root, levelToColumnNodesMap);
      this.formatColumnNodes(
        matrix.columns.root,
        levelToColumnNodesMap,
        tbodyElement,
        matrix.valueSources
      );
    } else {
      // ➕ Manually render value headers when only rows + values are used
      const trElement = document.createElement("tr");

      // Corner cell for row header
      const thCorner = document.createElement("th");
      thCorner.style.textAlign = "left";
      trElement.appendChild(thCorner);
      thCorner.style.color = "#fff";

      let rowFieldName: string = "";
      if (
        matrix.rows &&
        matrix.rows.levels &&
        matrix.rows.levels.length > 0 &&
        matrix.rows.levels[0].sources &&
        matrix.rows.levels[0].sources.length > 0
      ) {
        rowFieldName = matrix.rows.levels[0].sources[0].displayName;
      }

      console.log("rowFieldName", rowFieldName);
      console.log("matrix", matrix);

      // ✅ Add row field name in top-left corner
      thCorner.style.textAlign = "center";
      thCorner.style.boxShadow = "inset 0 -1px 0 0 #fff";
      thCorner.textContent = rowFieldName;
      thCorner.style.color = "#fff";
      trElement.appendChild(thCorner);

      // Each value field becomes a column header
      matrix.valueSources?.forEach((vs) => {
        const thElement = document.createElement("th");
        thElement.textContent = vs.displayName;
        thElement.style.color = "#fff";
        thElement.style.textAlign = "center";
        thElement.style.boxShadow =
          "inset 1px 0 0 0 #fff, inset 0 -1px 0 0 #fff";
        trElement.appendChild(thElement);
      });

      tbodyElement.appendChild(trElement);
    }

    this.formatRowNodes(matrix.rows.root, tbodyElement, matrix);

    tableElement.appendChild(tbodyElement);
    htmlElement.appendChild(tableElement);

    return htmlElement;
  }

  private static countColumnNodeLeaves(
    root,
    levelToColumnNodesMap: any[][]
  ): number {
    if (!(typeof root.level === "undefined" || root.level === null)) {
      if (!levelToColumnNodesMap[root.level]) {
        levelToColumnNodesMap[root.level] = [root];
      } else {
        levelToColumnNodesMap[root.level].push(root);
      }
    }
    let leafCount;
    if (root.isSubtotal || !root.children) {
      return (leafCount = 1);
    } else {
      leafCount = 0;
      for (const child of root.children) {
        leafCount += MatrixDataviewHtmlFormatter.countColumnNodeLeaves(
          child,
          levelToColumnNodesMap
        );
      }
    }
    // Store the leaf count on the node and return it
    root.leafCount = leafCount;
    return leafCount;
  }

  private static formatColumnNodes(
    root: any,
    levelToColumnNodesMap: any[][],
    topElement: HTMLElement,
    valueSources?: powerbi.DataViewMetadataColumn[]
  ) {
    // Create table rows for each level of column headers
    for (let level = 0; level < levelToColumnNodesMap.length; level++) {
      const levelNodes = levelToColumnNodesMap[level];
      const trElement = document.createElement("tr");

      // Add an empty top-left corner cell for row headers alignment
      const thCorner = document.createElement("th");
      thCorner.style.textAlign = "left";
      thCorner.style.color = "#fff";
      trElement.appendChild(thCorner);

      // Add column header cells for this level
      for (let i = 0; i < levelNodes.length; i++) {
        const node = levelNodes[i];
        const thElement = document.createElement("th");
        thElement.colSpan = node.leafCount;
        const headerText = node.isSubtotal
          ? "Totals"
          : node.value != null
          ? String(node.value)
          : (valueSources && valueSources[i] && valueSources[i].displayName) ||
            "Value";
        const textNode = document.createTextNode(headerText);
        thElement.appendChild(textNode);
        thElement.style.textAlign = "left";
        trElement.appendChild(thElement);
      }
      topElement.appendChild(trElement);
    }
  }

  private static formatRowNodes(
    root,
    topElement: HTMLElement,
    matrix: powerbi.DataViewMatrix
  ) {
    if (!(typeof root.level === "undefined" || root.level === null)) {
      const trElement = document.createElement("tr");

      const isEvenRow = MatrixDataviewHtmlFormatter.rowIndex % 2 === 0;
      const rowBackground = isEvenRow ? "#111111" : "#1a1a1a";
      MatrixDataviewHtmlFormatter.rowIndex++;

      const thElement = document.createElement("th");
      thElement.style.textAlign = "left";
      thElement.style.padding = "6px 10px";
      thElement.style.backgroundColor = rowBackground;
      thElement.style.color = "#fff";

      let headerText = "".padStart(root.level * 4, "\u00A0");
      headerText += root.isSubtotal ? "Totals" : String(root.value);
      thElement.textContent = headerText;
      trElement.appendChild(thElement);

      // Data cells for each measure/column intersection
      if (root.values) {
        for (
          let i = 0;
          !(typeof root.values[i] === "undefined" || root.values[i] === null);
          i++
        ) {
          const value = root.values[i]?.value;
          const td = document.createElement("td");

          td.style.backgroundColor = rowBackground;
          td.style.padding = "6px 10px";
          td.style.whiteSpace = "nowrap";
          td.style.fontSize = "13px";
          td.style.color = "#fff";
          td.style.textAlign = "center";

          td.style.boxShadow = "inset 1px 0 0 0 #fff";

          let finalColor: string | undefined = undefined;

          // Traffic light color map
          const colorMap: Record<string, string> = {
            red: "#d64150",
            green: "#4caf50",
            yellow: "#ffc107",
          };

          let commentColor: string | undefined = undefined;

          if (value != null && !isNaN(value)) {
            const valNum = Number(value);
            const formatted = `${valNum.toFixed(1)} €`;

            const rowLabel = String(root.value); // row name
            const columnName =
              matrix.valueSources?.[i]?.displayName || `Column${i}`;
            const filterId = md5(`${rowLabel}_${columnName}`);
            td.setAttribute("data-filter-id", filterId);
            const allComments = [...(this.commentMap.get(filterId) || [])];

            // ✅ Sort by createdAt descending (most recent first)
            allComments.sort((a, b) => {
              const aDate = new Date(a.createdAt || "").getTime();
              const bDate = new Date(b.createdAt || "").getTime();
              return bDate - aDate;
            });

            // Check if comment exists in commentMap
            if (this.commentMap.has(filterId)) {
              const latestComment = allComments[0];
              if (latestComment?.color) {
                switch (latestComment.color) {
                  case "red":
                    td.style.backgroundColor = "#ff4d4f";
                    break;
                  case "green":
                    td.style.backgroundColor = "#73d13d";
                    break;
                  case "yellow":
                    td.style.backgroundColor = "#faad14";
                    break;
                  case "default":
                  default:
                    td.style.backgroundColor = rowBackground;
                    break;
                }
              }

              td.title = allComments
                .map((c) => `💬 ${c.comment}`)
                .join("\n──────────────\n");

              // Pick the last comment's color for display
              const lastColor = allComments[0]?.color || "default";

              if (lastColor !== "default" && colorMap[lastColor]) {
                finalColor = colorMap[lastColor]; // 🟢 Comment traffic light color
                commentColor = lastColor;
              }

              // td.style.outline = "1px dashed yellow";

              // ⚠️ Add alert icon only if there’s at least one non-empty comment
              const hasNonEmptyComment = allComments.some(
                (c) => c.comment.trim() !== ""
              );
              if (hasNonEmptyComment) {
                const alertIcon = document.createElement("img");
                alertIcon.src =
                  "data:image/svg+xml;base64,PD94bWwgdmVyc2lvbj0iMS4wIiBzdGFuZGFsb25lPSJubyI/Pgo8IURPQ1RZUEUgc3ZnIFBVQkxJQyAiLS8vVzNDLy9EVEQgU1ZHIDIwMDEwOTA0Ly9FTiIKICJodHRwOi8vd3d3LnczLm9yZy9UUi8yMDAxL1JFQy1TVkctMjAwMTA5MDQvRFREL3N2ZzEwLmR0ZCI+CjxzdmcgdmVyc2lvbj0iMS4wIiB4bWxucz0iaHR0cDovL3d3dy53My5vcmcvMjAwMC9zdmciCiB3aWR0aD0iODYwLjAwMDAwMHB0IiBoZWlnaHQ9Ijc4MC4wMDAwMDBwdCIgdmlld0JveD0iMCAwIDg2MC4wMDAwMDAgNzgwLjAwMDAwMCIKIHByZXNlcnZlQXNwZWN0UmF0aW89InhNaWRZTWlkIG1lZXQiPgoKPGcgdHJhbnNmb3JtPSJ0cmFuc2xhdGUoMC4wMDAwMDAsNzgwLjAwMDAwMCkgc2NhbGUoMC4xMDAwMDAsLTAuMTAwMDAwKSIKZmlsbD0iIzAwMDAwMCIgc3Ryb2tlPSJub25lIj4KPHBhdGggZD0iTTQxNDIgNzM3NSBjLTE5MyAtNDQgLTM5MCAtMTgyIC01MTcgLTM2MCAtNjUgLTkxIC0yMDIgLTMxNyAtMTYyOAotMjY3NSAtMTUwNiAtMjQ5MCAtMTYxMSAtMjY2NyAtMTY0MCAtMjc1NyAtMzIgLTk2IC01NyAtMjI4IC01NyAtMjk4IDAgLTE4Ngo4MCAtMzc2IDIxNSAtNTExIDg4IC04OCAxOTggLTE1NCAzMjUgLTE5NSA4NSAtMjggMTE1IC0zMiAyODEgLTQwIDEwOCAtNgoxNTEyIC04IDMzNDAgLTYgMzE1MiAzIDMxNTQgMyAzMjI5IDIzIDMzMSA5MiA1MzggMjk5IDU5NSA1OTcgMTkgMTAwIDE5IDE2NAotMSAyNjggLTM1IDE4NCAtNDcgMjA4IC0zNTggNzI2IC00MTggNjk3IC0xMjg3IDIxMzcgLTE4OTEgMzEzMyAtNzkyIDEzMDcKLTk5NCAxNjMzIC0xMDk3IDE3NzMgLTEwMCAxMzUgLTI1MCAyNDYgLTQwNiAzMDAgLTc1IDI2IC0xMDMgMzEgLTIwNyAzMyAtNzgKMiAtMTQyIC0yIC0xODMgLTExeiBtMjg0IC01MTggYzkxIC00NiA2OSAtMTQgNjYyIC05ODcgMTAyMyAtMTY3OSAyMjEzIC0zNjQ0CjI1NjcgLTQyMzkgMTM0IC0yMjQgMTYxIC0zMDQgMTM2IC0zOTggLTE1IC01MyAtNzAgLTEzMyAtMTA3IC0xNTUgLTc0IC00MwotNzEgLTQzIC0zMzg0IC00MyAtMzMxMyAwIC0zMzEwIDAgLTMzODQgNDMgLTM3IDIyIC05MiAxMDIgLTEwNyAxNTUgLTI1IDk0IDIKMTc0IDEzNiAzOTggMzE4IDUzNSAxNzY0IDI5MjIgMjUwNyA0MTM5IDEyOSAyMTIgMzE1IDUxOCA0MTQgNjgwIDE5NyAzMjQgMjQ0CjM4MyAzMzQgNDE5IDcyIDI4IDE1NSAyNCAyMjYgLTEyeiIvPgo8cGF0aCBkPSJNNDIyMCA1Mzc0IGMtNjkgLTI0IC0xMTAgLTYxIC0xMzkgLTEyMyBsLTI2IC01NiAtMyAtODY4IGMtMiAtNTE3IDEKLTkyMyA3IC0xMDA0IDExIC0xNTkgMzAgLTIwNiA5OSAtMjU3IDQyIC0zMCA0OCAtMzEgMTQyIC0zMSA5NCAwIDEwMCAxIDE0MgozMSA2NyA0OSA4NyA5OSA5OCAyNDAgNSA2NSAxMCA1MTkgMTAgMTAwOCBsMCA4ODkgLTIyIDM2IGMtNzMgMTIzIC0xOTQgMTc2Ci0zMDggMTM1eiIvPgo8cGF0aCBkPSJNNDIxNyAyMzkwIGMtMTM5IC0zNSAtMjY0IC0xNzIgLTI4NyAtMzE0IC0zNyAtMjI5IDIwMyAtNDY1IDQyNgotNDE4IDI0NyA1MSAzODcgMzEwIDI4MSA1MTggLTMzIDY1IC0xMjYgMTU1IC0xOTQgMTg5IC02MSAyOSAtMTYyIDQxIC0yMjYgMjV6Ii8+CjwvZz4KPC9zdmc+Cg==";

                alertIcon.alt = "!";
                alertIcon.style.position = "absolute";
                alertIcon.style.top = "2px";
                alertIcon.style.right = "4px";
                alertIcon.style.width = "14px";
                alertIcon.style.height = "14px";

                // ❗ Choose alert color by logic
                if (lastColor === "green" || lastColor === "red") {
                  alertIcon.style.filter = "drop-shadow(0 0 1px yellow)";
                } else if (lastColor === "yellow") {
                  alertIcon.style.filter = "drop-shadow(0 0 1px red)";
                }

                td.style.position = "relative";
                td.appendChild(alertIcon);
              }
            }

            const bar = document.createElement("div");
            bar.textContent = formatted;
            bar.style.display = "inline-block";
            bar.style.padding = "2px 6px";
            bar.style.borderRadius = "2px";
            bar.style.minWidth = "60px";

            // ✅ Apply red bar only for configured columns and when value > threshold
            // if (
            //   this.config.redBarColumnsIndex.includes(i) &&
            //   valNum > this.config.redBarThreshold
            // ) {
            //   bar.style.backgroundColor = "#d64150";
            //   td.style.backgroundColor = "#d64150";
            // }

            // 🚨 If no traffic light color is set (or it's default), apply red bar rule
            if (
              !finalColor &&
              this.config.redBarColumnsIndex.includes(i) &&
              valNum > this.config.redBarThreshold
            ) {
              finalColor = "#d64150";
            }

            // ✅ Apply background color
            if (finalColor) {
              td.style.backgroundColor = finalColor;
              bar.style.backgroundColor = finalColor;
            }

            td.appendChild(bar);
            if (
              !root.isSubtotal &&
              this.config.redBarColumnsIndex.includes(i)
            ) {
              td.setAttribute("data-clickable", "true");
              td.style.cursor = "pointer"; // ✅ Enable pointer only for redBarColumnsIndex
            } else {
              td.style.cursor = "default"; // ❌ Force non-clickable cells to not show pointer
            }
          }
          trElement.appendChild(td);
        }
      }

      topElement.appendChild(trElement);
    }

    // Recurse into children (sub-rows) if present
    if (root.children) {
      for (const child of root.children) {
        MatrixDataviewHtmlFormatter.formatRowNodes(child, topElement, matrix);
      }
    }
  }
}
