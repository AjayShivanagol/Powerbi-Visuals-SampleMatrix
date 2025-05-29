"use strict";
import powerbi from "powerbi-visuals-api";
import md5 from "md5";
import { DataViewObjectPropertyReference, Selector } from "./common";
import { MatrixDataviewHtmlFormatter } from "./matrixDataviewHtmlFormatter";
import { ObjectEnumerationBuilder } from "./objectEnumerationBuilder";
import { SubtotalProperties } from "./subtotalProperties";
import IVisual = powerbi.extensibility.visual.IVisual;
import VisualConstructorOptions = powerbi.extensibility.visual.VisualConstructorOptions;
import VisualUpdateOptions = powerbi.extensibility.visual.VisualUpdateOptions;
import DataView = powerbi.DataView;
import EnumerateVisualObjectInstancesOptions = powerbi.EnumerateVisualObjectInstancesOptions;
import VisualObjectInstance = powerbi.VisualObjectInstance;
import VisualObjectInstanceEnumerationObject = powerbi.VisualObjectInstanceEnumerationObject;
import DataViewObjects = powerbi.DataViewObjects;
import DataViewObject = powerbi.DataViewObject;
import DataViewHierarchyLevel = powerbi.DataViewHierarchyLevel;
import DataViewMatrixNodeValue = powerbi.DataViewMatrixNodeValue;

export class Visual implements IVisual {
  private target: HTMLElement;
  private dataView: DataView;
  private refreshInterval: number | undefined;
  private upnValue: string = "";

  constructor(options: VisualConstructorOptions) {
    console.log("Visual constructor", options);
    this.target = options.element;
    options.element.style.overflow = "auto";
    this.target.style.overflow = "auto";
    this.target.style.height = "100%";
  }

  private generateUUID(): string {
    return "xxxxxxxx-xxxx-4xxx-yxxx-xxxxxxxxxxxx".replace(/[xy]/g, (c) => {
      const r = (Math.random() * 16) | 0;
      const v = c === "x" ? r : (r & 0x3) | 0x8;
      return v.toString(16);
    });
  }

  private refreshCommentsAndRedraw() {
    fetch(MatrixDataviewHtmlFormatter.config.commentApiUrl)
      .then((res) => res.json())
      .then((comments) => {
        console.log("reponse comments", comments);
        MatrixDataviewHtmlFormatter.setCommentMap(comments);
        const updated = MatrixDataviewHtmlFormatter.formatDataViewMatrix(
          this.dataView.matrix
        );
        this.target.innerHTML = "";
        this.target.appendChild(updated);

        const clickableCells =
          this.target.querySelectorAll("td[data-clickable]");
        clickableCells.forEach((cell) => {
          const colIndex = Array.from(cell.parentElement!.children).indexOf(
            cell
          );
          if (
            MatrixDataviewHtmlFormatter.config.redBarColumnsIndex.includes(
              colIndex - 1
            )
          ) {
            cell.addEventListener("click", () => {
              this.openCommentPopup(cell as HTMLElement);
            });
          }
        });
      });
  }

  public update(options: VisualUpdateOptions) {
    if (!options) {
      return;
    }

    const matrix = options.dataViews![0].matrix;
    const upnIdx = matrix.valueSources.findIndex(
      (vs) => vs.roles?.UserPrincipalName
    );

    // pull out both the cell’s raw value and its column‐index “key”
    const upnNode = matrix.rows.root.children[0].values[
      upnIdx
    ] as DataViewMatrixNodeValue;
    const { value: rawUpn = "", valueSourceIndex: key } = upnNode;

    // coerce to string so we can call .split()
    const rawUpnStr = String(rawUpn);
    this.upnValue = rawUpnStr.split("@")[0].toUpperCase();

    console.log("UPN:", this.upnValue, "key:", key);

    if (options.type & powerbi.VisualUpdateType.Data) {
      if (
        !options.dataViews ||
        !options.dataViews[0] ||
        !options.dataViews[0].matrix ||
        !options.dataViews[0].matrix.rows ||
        !options.dataViews[0].matrix.rows.root ||
        !options.dataViews[0].matrix.rows.root.children ||
        !options.dataViews[0].matrix.rows.root.children.length ||
        !options.dataViews[0].matrix.columns ||
        !options.dataViews[0].matrix.columns.root ||
        !options.dataViews[0].matrix.columns.root.children ||
        !options.dataViews[0].matrix.columns.root.children.length
      ) {
        this.dataView = undefined;
        return;
      }

      this.dataView = options.dataViews[0];

      while (this.target.firstChild) {
        this.target.removeChild(this.target.firstChild);
      }

      // this.target.appendChild(
      //   MatrixDataviewHtmlFormatter.formatDataViewMatrix(
      //     options.dataViews[0].matrix
      //   )
      // );
      // Create the matrix table HTML from the data view and append it
      // Append the new matrix table HTML

      const objects = options.dataViews?.[0]?.metadata?.objects;
      const settings = objects?.visualSettings;

      const redBarThreshold = (settings?.redBarThreshold as number) ?? 0;
      const redBarColumnsRaw = (
        (settings?.redBarColumnsIndex as string) ?? ""
      ).split(",");
      const redBarColumnsIndex = redBarColumnsRaw
        .map((c) => parseInt(c.trim()))
        .filter((n) => !isNaN(n));
      const commentApiUrl =
        (settings?.commentApiUrl as string) ??
        "https://default-url/api/comment";

      // Set config to formatter
      MatrixDataviewHtmlFormatter.setConfig({
        redBarThreshold,
        redBarColumnsIndex,
        commentApiUrl,
      });

      const tableElement = MatrixDataviewHtmlFormatter.formatDataViewMatrix(
        options.dataViews[0].matrix
      );
      this.target.appendChild(tableElement);

      // Attach click handlers to each cell marked as clickable
      const clickableCells = this.target.querySelectorAll("td[data-clickable]");
      clickableCells.forEach((cell) => {
        const colIndex = Array.from(cell.parentElement!.children).indexOf(cell);

        // Only attach listener if the column is in the configured red bar list
        if (
          MatrixDataviewHtmlFormatter.config.redBarColumnsIndex.includes(
            colIndex - 1
          )
        ) {
          cell.addEventListener("click", () => {
            this.openCommentPopup(cell as HTMLElement);
          });
        }
      });
    }

    if (!this.refreshInterval) {
      this.refreshInterval = setInterval(() => {
        const popupExists = this.target.querySelector(".comment-popup");
        if (popupExists) {
          console.log("Skipping redraw: comment popup is open.");
          return; // ❌ Don't redraw matrix if popup is active
        }
        this.refreshCommentsAndRedraw();
      }, 4000);
    }
  }

  // Create and display the comment popup for the clicked cell
  private openCommentPopup(cell: HTMLElement): void {
    // Remove any existing popup to only show one at a time
    const existingPopup = this.target.querySelector(
      ".comment-popup"
    ) as HTMLElement;
    if (existingPopup) {
      this.target.removeChild(existingPopup);
    }

    // Find column index
    // ✅ Correct header detection using valueSources
    let columnHeaderText = "Unknown Column";
    const cellIndex =
      Array.from(cell.parentElement!.children).indexOf(cell) - 1; // exclude row header
    const valueSources = this.dataView?.matrix?.valueSources;

    if (
      valueSources &&
      cellIndex >= 0 &&
      cellIndex < valueSources.length &&
      valueSources[cellIndex].displayName
    ) {
      columnHeaderText = valueSources[cellIndex].displayName;
    }

    const cellText = cell.textContent?.trim() || "Empty";

    const rowLabel =
      cell.parentElement?.querySelector("th")?.textContent?.trim() || "Row";
    const filterId = md5(`${rowLabel}_${columnHeaderText}`);

    // Create popup container
    const popup = document.createElement("div");
    popup.className = "comment-popup";
    // Basic styling for the popup (to ensure visibility)
    popup.style.position = "fixed";
    popup.style.top = "50%";
    popup.style.left = "50%";
    popup.style.transform = "translate(-50%, -50%)";
    popup.style.background = "#fff";
    popup.style.padding = "16px";
    popup.style.border = "1px solid #aaa";
    popup.style.zIndex = "1000";
    popup.style.width = "450px";
    popup.style.maxHeight = "400px";
    popup.style.overflowY = "auto";
    popup.style.borderRadius = "8px";

    // Add header text
    const headerDiv = document.createElement("div");
    headerDiv.style.fontWeight = "bold";
    headerDiv.style.marginBottom = "10px";
    headerDiv.textContent = `Comments for ${columnHeaderText} - ${rowLabel}`;
    popup.appendChild(headerDiv);

    // Comments list
    const allComments: {
      id: string;
      user: string;
      column: string;
      comment: string;
      filterId: string;
      createdAt?: string;
      modifiedAt?: string;
      color?: string;
    }[] = MatrixDataviewHtmlFormatter["commentMap"].get(filterId) || [];

    // Sort all Comments by createdAt descending (most recent first)
    allComments.sort((a, b) => {
      const aDate = new Date(a.createdAt || "").getTime();
      const bDate = new Date(b.createdAt || "").getTime();
      return aDate - bDate; // Sorts ascending
    });

    const lastColor =
      allComments.length > 0
        ? allComments[allComments.length - 1]?.color || "default"
        : "default";
    console.log("lastColor", lastColor);

    const commentList = document.createElement("div");
    commentList.style.marginBottom = "10px";

    // 4) filter out empties for display
    const displayComments = allComments.filter((c) => c.comment.trim() !== "");

    displayComments.forEach((commentObj, index) => {
      const commentDiv = document.createElement("div");
      commentDiv.style.display = "flex";
      commentDiv.style.flexDirection = "column";
      commentDiv.style.gap = "4px";
      commentDiv.style.padding = "10px 0";
      commentDiv.style.borderBottom = "1px solid #ccc";

      const topRow = document.createElement("div");
      topRow.style.display = "flex";
      topRow.style.alignItems = "center"; // center‐vertically
      topRow.style.gap = "8px";

      // --- NEW: show the commenting user ---
      const avatar = document.createElement("div");
      const initials = commentObj.user
        .split(" ")
        .map((w) => w[0])
        .join("")
        .slice(0, 2)
        .toUpperCase();
      avatar.textContent = initials;
      avatar.title = commentObj.user;
      Object.assign(avatar.style, {
        width: "32px",
        height: "32px",
        borderRadius: "50%",
        backgroundColor: "#555",
        color: "#fff",
        fontSize: "14px",
        textAlign: "center",
        lineHeight: "32px",
        userSelect: "none",
      });

      const commentContent = document.createElement("div");
      commentContent.textContent = commentObj.comment;
      commentContent.style.whiteSpace = "pre-wrap";
      commentContent.style.flex = "1";

      const iconContainer = document.createElement("div");
      iconContainer.style.display = "flex";
      iconContainer.style.gap = "10px";

      const timestamp = document.createElement("div");
      timestamp.style.fontSize = "8px";
      timestamp.style.color = "#888";
      timestamp.style.display = "flex";
      timestamp.style.alignSelf = "flex-end";
      timestamp.style.alignItems = "center";
      timestamp.style.gap = "4px";
      timestamp.style.marginLeft = "auto";

      const buttonRow = document.createElement("div");
      buttonRow.style.display = "flex";
      buttonRow.style.gap = "10px";
      buttonRow.style.alignItems = "center";

      // Edit icon (SVG)
      const editBtn = document.createElement("img");
      editBtn.src =
        "data:image/svg+xml;base64,PD94bWwgdmVyc2lvbj0iMS4wIiA/PjwhRE9DVFlQRSBzdmcgIFBVQkxJQyAnLS8vVzNDLy9EVEQgU1ZHIDEuMS8vRU4nICAnaHR0cDovL3d3dy53My5vcmcvR3JhcGhpY3MvU1ZHLzEuMS9EVEQvc3ZnMTEuZHRkJz48IS0tIFVwbG9hZGVkIHRvOiBTVkcgUmVwbywgd3d3LnN2Z3JlcG8uY29tLCBHZW5lcmF0b3I6IFNWRyBSZXBvIE1peGVyIFRvb2xzIC0tPgo8c3ZnIGZpbGw9IiMwMDAwMDAiIHdpZHRoPSI4MDBweCIgaGVpZ2h0PSI4MDBweCIgdmlld0JveD0iMCAwIDMyIDMyIiBzdHlsZT0iZmlsbC1ydWxlOmV2ZW5vZGQ7Y2xpcC1ydWxlOmV2ZW5vZGQ7c3Ryb2tlLWxpbmVqb2luOnJvdW5kO3N0cm9rZS1taXRlcmxpbWl0OjI7IiB2ZXJzaW9uPSIxLjEiIHhtbDpzcGFjZT0icHJlc2VydmUiIHhtbG5zPSJodHRwOi8vd3d3LnczLm9yZy8yMDAwL3N2ZyIgeG1sbnM6c2VyaWY9Imh0dHA6Ly93d3cuc2VyaWYuY29tLyIgeG1sbnM6eGxpbms9Imh0dHA6Ly93d3cudzMub3JnLzE5OTkveGxpbmsiPjxwYXRoIGQ9Ik0xMi45NjUsNS40NjJjMCwtMCAtMi41ODQsMC4wMDQgLTQuOTc5LDAuMDA4Yy0zLjAzNCwwLjAwNiAtNS40OSwyLjQ2NyAtNS40OSw1LjVsMCwxMy4wM2MwLDEuNDU5IDAuNTc5LDIuODU4IDEuNjExLDMuODg5YzEuMDMxLDEuMDMyIDIuNDMsMS42MTEgMy44ODksMS42MTFsMTMuMDAzLDBjMy4wMzgsLTAgNS41LC0yLjQ2MiA1LjUsLTUuNWMwLC0yLjQwNSAwLC01LjAwNCAwLC01LjAwNGMwLC0wLjgyOCAtMC42NzIsLTEuNSAtMS41LC0xLjVjLTAuODI3LC0wIC0xLjUsMC42NzIgLTEuNSwxLjVsMCw1LjAwNGMwLDEuMzgxIC0xLjExOSwyLjUgLTIuNSwyLjVsLTEzLjAwMywwYy0wLjY2MywtMCAtMS4yOTksLTAuMjYzIC0xLjc2OCwtMC43MzJjLTAuNDY5LC0wLjQ2OSAtMC43MzIsLTEuMTA1IC0wLjczMiwtMS43NjhsMCwtMTMuMDNjMCwtMS4zNzkgMS4xMTcsLTIuNDk3IDIuNDk2LC0yLjVjMi4zOTQsLTAuMDA0IDQuOTc5LC0wLjAwOCA0Ljk3OSwtMC4wMDhjMC44MjgsLTAuMDAyIDEuNDk4LC0wLjY3NSAxLjQ5NywtMS41MDNjLTAuMDAxLC0wLjgyOCAtMC42NzUsLTEuNDk5IC0xLjUwMywtMS40OTdaIi8+PHBhdGggZD0iTTIwLjA0Niw2LjQxMWwtNi44NDUsNi44NDZjLTAuMTM3LDAuMTM3IC0wLjIzMiwwLjMxMSAtMC4yNzEsMC41MDFsLTEuMDgxLDUuMTUyYy0wLjA2OSwwLjMyOSAwLjAzMiwwLjY3MSAwLjI2OCwwLjkwOWMwLjIzNywwLjIzOSAwLjU3NywwLjM0MyAwLjkwNywwLjI3N2w1LjE5NCwtMS4wMzhjMC4xOTMsLTAuMDM5IDAuMzcxLC0wLjEzNCAwLjUxMSwtMC4yNzRsNi44NDUsLTYuODQ1bC01LjUyOCwtNS41MjhabTEuNDE1LC0xLjQxNGw1LjUyNyw1LjUyOGwxLjExMiwtMS4xMTFjMS41MjYsLTEuNTI3IDEuNTI2LC00LjAwMSAtMCwtNS41MjdjLTAuMDAxLC0wIC0wLjAwMSwtMC4wMDEgLTAuMDAxLC0wLjAwMWMtMS41MjcsLTEuNTI2IC00LjAwMSwtMS41MjYgLTUuNTI3LC0wbC0xLjExMSwxLjExMVoiLz48ZyBpZD0iSWNvbiIvPjwvc3ZnPg==";
      editBtn.alt = "Edit";
      editBtn.title = "Edit";
      editBtn.style.width = "16px";
      editBtn.style.height = "16px";
      editBtn.style.cursor = "pointer";
      editBtn.onclick = () => {
        textarea.value = commentObj.comment;
        textarea.dataset.editIndex = index.toString();
        textarea.dataset.commentId = commentObj.id;
      };

      // Delete icon (emoji or SVG)
      const deleteBtn = document.createElement("img");
      deleteBtn.src =
        "data:image/svg+xml;base64,PD94bWwgdmVyc2lvbj0iMS4wIiBlbmNvZGluZz0iaXNvLTg4NTktMSI/Pg0KPCEtLSBVcGxvYWRlZCB0bzogU1ZHIFJlcG8sIHd3dy5zdmdyZXBvLmNvbSwgR2VuZXJhdG9yOiBTVkcgUmVwbyBNaXhlciBUb29scyAtLT4NCjwhRE9DVFlQRSBzdmcgUFVCTElDICItLy9XM0MvL0RURCBTVkcgMS4xLy9FTiIgImh0dHA6Ly93d3cudzMub3JnL0dyYXBoaWNzL1NWRy8xLjEvRFREL3N2ZzExLmR0ZCI+DQo8c3ZnIGZpbGw9IiMwMDAwMDAiIHZlcnNpb249IjEuMSIgaWQ9IkNhcGFfMSIgeG1sbnM9Imh0dHA6Ly93d3cudzMub3JnLzIwMDAvc3ZnIiB4bWxuczp4bGluaz0iaHR0cDovL3d3dy53My5vcmcvMTk5OS94bGluayIgDQoJIHdpZHRoPSI4MDBweCIgaGVpZ2h0PSI4MDBweCIgdmlld0JveD0iMCAwIDQ4Mi40MjggNDgyLjQyOSINCgkgeG1sOnNwYWNlPSJwcmVzZXJ2ZSI+DQo8Zz4NCgk8Zz4NCgkJPHBhdGggZD0iTTM4MS4xNjMsNTcuNzk5aC03NS4wOTRDMzAyLjMyMywyNS4zMTYsMjc0LjY4NiwwLDI0MS4yMTQsMGMtMzMuNDcxLDAtNjEuMTA0LDI1LjMxNS02NC44NSw1Ny43OTloLTc1LjA5OA0KCQkJYy0zMC4zOSwwLTU1LjExMSwyNC43MjgtNTUuMTExLDU1LjExN3YyLjgyOGMwLDIzLjIyMywxNC40Niw0My4xLDM0LjgzLDUxLjE5OXYyNjAuMzY5YzAsMzAuMzksMjQuNzI0LDU1LjExNyw1NS4xMTIsNTUuMTE3DQoJCQloMjEwLjIzNmMzMC4zODksMCw1NS4xMTEtMjQuNzI5LDU1LjExMS01NS4xMTdWMTY2Ljk0NGMyMC4zNjktOC4xLDM0LjgzLTI3Ljk3NywzNC44My01MS4xOTl2LTIuODI4DQoJCQlDNDM2LjI3NCw4Mi41MjcsNDExLjU1MSw1Ny43OTksMzgxLjE2Myw1Ny43OTl6IE0yNDEuMjE0LDI2LjEzOWMxOS4wMzcsMCwzNC45MjcsMTMuNjQ1LDM4LjQ0MywzMS42NmgtNzYuODc5DQoJCQlDMjA2LjI5MywzOS43ODMsMjIyLjE4NCwyNi4xMzksMjQxLjIxNCwyNi4xMzl6IE0zNzUuMzA1LDQyNy4zMTJjMCwxNS45NzgtMTMsMjguOTc5LTI4Ljk3MywyOC45NzlIMTM2LjA5Ng0KCQkJYy0xNS45NzMsMC0yOC45NzMtMTMuMDAyLTI4Ljk3My0yOC45NzlWMTcwLjg2MWgyNjguMTgyVjQyNy4zMTJ6IE00MTAuMTM1LDExNS43NDRjMCwxNS45NzgtMTMsMjguOTc5LTI4Ljk3MywyOC45NzlIMTAxLjI2Ng0KCQkJYy0xNS45NzMsMC0yOC45NzMtMTMuMDAxLTI4Ljk3My0yOC45Nzl2LTIuODI4YzAtMTUuOTc4LDEzLTI4Ljk3OSwyOC45NzMtMjguOTc5aDI3OS44OTdjMTUuOTczLDAsMjguOTczLDEzLjAwMSwyOC45NzMsMjguOTc5DQoJCQlWMTE1Ljc0NHoiLz4NCgkJPHBhdGggZD0iTTE3MS4xNDQsNDIyLjg2M2M3LjIxOCwwLDEzLjA2OS01Ljg1MywxMy4wNjktMTMuMDY4VjI2Mi42NDFjMC03LjIxNi01Ljg1Mi0xMy4wNy0xMy4wNjktMTMuMDcNCgkJCWMtNy4yMTcsMC0xMy4wNjksNS44NTQtMTMuMDY5LDEzLjA3djE0Ny4xNTRDMTU4LjA3NCw0MTcuMDEyLDE2My45MjYsNDIyLjg2MywxNzEuMTQ0LDQyMi44NjN6Ii8+DQoJCTxwYXRoIGQ9Ik0yNDEuMjE0LDQyMi44NjNjNy4yMTgsMCwxMy4wNy01Ljg1MywxMy4wNy0xMy4wNjhWMjYyLjY0MWMwLTcuMjE2LTUuODU0LTEzLjA3LTEzLjA3LTEzLjA3DQoJCQljLTcuMjE3LDAtMTMuMDY5LDUuODU0LTEzLjA2OSwxMy4wN3YxNDcuMTU0QzIyOC4xNDUsNDE3LjAxMiwyMzMuOTk2LDQyMi44NjMsMjQxLjIxNCw0MjIuODYzeiIvPg0KCQk8cGF0aCBkPSJNMzExLjI4NCw0MjIuODYzYzcuMjE3LDAsMTMuMDY4LTUuODUzLDEzLjA2OC0xMy4wNjhWMjYyLjY0MWMwLTcuMjE2LTUuODUyLTEzLjA3LTEzLjA2OC0xMy4wNw0KCQkJYy03LjIxOSwwLTEzLjA3LDUuODU0LTEzLjA3LDEzLjA3djE0Ny4xNTRDMjk4LjIxMyw0MTcuMDEyLDMwNC4wNjcsNDIyLjg2MywzMTEuMjg0LDQyMi44NjN6Ii8+DQoJPC9nPg0KPC9nPg0KPC9zdmc+";
      deleteBtn.alt = "Delete";
      deleteBtn.title = "Delete";
      deleteBtn.style.width = "16px";
      deleteBtn.style.height = "16px";
      deleteBtn.style.cursor = "pointer";
      deleteBtn.onclick = () => {
        fetch(
          `${MatrixDataviewHtmlFormatter.config.commentApiUrl}/${filterId}`,
          {
            method: "DELETE",
            headers: { "Content-Type": "application/json" },
            body: JSON.stringify({ comment: commentObj.id }),
          }
        )
          .then(() => this.refreshCommentsAndRedraw())
          .catch((err) => console.error("Failed to delete comment", err));
        this.target.removeChild(popup);
      };

      // Only show edit/delete if the comment is by the current user
      if (commentObj.user?.toUpperCase() === this.upnValue) {
        iconContainer.appendChild(editBtn);
        iconContainer.appendChild(deleteBtn);
      }
      topRow.appendChild(avatar);
      topRow.appendChild(commentContent);
      topRow.appendChild(iconContainer);
      commentDiv.appendChild(topRow);

      if (commentObj.createdAt) {
        const time = new Date(commentObj.createdAt);
        const formattedDate = `${String(time.getDate()).padStart(
          2,
          "0"
        )}-${time.toLocaleString("default", {
          month: "short",
        })}-${time.getFullYear()} ${String(time.getHours()).padStart(
          2,
          "0"
        )}:${String(time.getMinutes()).padStart(2, "0")}`;

        const timestamp = document.createElement("div");
        timestamp.style.fontSize = "11px";
        timestamp.style.color = "#888";
        timestamp.style.display = "flex";
        timestamp.style.alignSelf = "flex-end";
        timestamp.style.alignItems = "center";
        timestamp.style.gap = "4px";

        const clockIcon = document.createElement("span");
        clockIcon.textContent = "🕒";
        const timeSpan = document.createElement("span");
        timeSpan.textContent = formattedDate;

        timestamp.appendChild(clockIcon);
        timestamp.appendChild(timeSpan);
        commentDiv.appendChild(timestamp);
      }
      commentList.appendChild(commentDiv);
    });

    popup.appendChild(commentList);

    // Color selection UI
    // 🔴🟢🟡 Color selection UI (Traffic Lights)
    const colorTitle = document.createElement("div");
    colorTitle.textContent = "Traffic Light Color:";
    colorTitle.style.fontWeight = "bold";
    colorTitle.style.marginBottom = "6px";
    popup.appendChild(colorTitle);

    const colorContainer = document.createElement("div");
    colorContainer.style.display = "flex";
    colorContainer.style.gap = "12px";
    colorContainer.style.marginBottom = "12px";

    // Map of color name to emoji or label for visual hint
    const colorOptions = [
      { value: "default", label: "Default" },
      { value: "red", label: "Red" },
      { value: "green", label: "Green" },
      { value: "yellow", label: "Yellow" },
    ];

    colorOptions.forEach(({ value, label }) => {
      const wrapper = document.createElement("label");
      wrapper.style.display = "flex";
      wrapper.style.alignItems = "center";
      wrapper.style.gap = "4px";
      wrapper.style.fontSize = "13px";

      const input = document.createElement("input");
      input.type = "radio";
      input.name = "commentColor";
      input.value = value;
      if (value === lastColor) input.checked = true;

      wrapper.appendChild(input);
      wrapper.appendChild(document.createTextNode(label));
      colorContainer.appendChild(wrapper);
    });

    popup.appendChild(colorContainer);

    // Textarea for new or edited comment
    const textarea = document.createElement("textarea");
    textarea.rows = 4;
    textarea.cols = 30;
    textarea.placeholder = "Add a comment...";
    popup.appendChild(textarea);
    popup.appendChild(document.createElement("br"));

    // Submit and Cancel buttons
    const submitBtn = document.createElement("button");
    submitBtn.textContent = "Submit";
    submitBtn.style.marginRight = "8px";
    submitBtn.style.background = "#0078d4";
    submitBtn.style.color = "#fff";
    submitBtn.style.border = "none";
    submitBtn.style.padding = "6px 12px";
    submitBtn.style.borderRadius = "4px";
    submitBtn.style.cursor = "pointer";
    popup.appendChild(submitBtn);

    const cancelBtn = document.createElement("button");
    cancelBtn.style.background = "#f3f3f3";
    cancelBtn.style.color = "#333";
    cancelBtn.style.border = "1px solid #ccc";
    cancelBtn.style.padding = "6px 12px";
    cancelBtn.style.borderRadius = "4px";
    cancelBtn.style.cursor = "pointer";
    cancelBtn.textContent = "Cancel";
    popup.appendChild(cancelBtn);

    cancelBtn.onclick = () => {
      if (popup.parentElement) this.target.removeChild(popup);
      textarea.dataset.commentId = "";
      textarea.dataset.editIndex = "";
    };

    submitBtn.onclick = () => {
      const commentText = textarea.value.trim();
      const selectedColor =
        (
          popup.querySelector(
            'input[name="commentColor"]:checked'
          ) as HTMLInputElement
        )?.value ?? "default";

      const editIndex = textarea.dataset.editIndex;
      const isEdit = textarea.dataset.commentId !== undefined;

      const previousColor =
        allComments[allComments.length - 1]?.color ?? "default";
      const isColorChanged = selectedColor !== previousColor;

      // 🔒 Prevent saving if no comment and no color change
      if (!commentText && !isColorChanged) return;

      console.log(
        "edit mode?",
        isEdit,
        "commentId:",
        textarea.dataset.commentId
      );

      if (isEdit) {
        const commentId = textarea.dataset.commentId;
        const index = parseInt(editIndex, 10);
        const existingComment = allComments[index];

        const selectedColor =
          (
            popup.querySelector(
              'input[name="commentColor"]:checked'
            ) as HTMLInputElement
          )?.value ?? "default";

        const updatedComment = {
          comment: commentText,
          modifiedAt: new Date().toISOString(),
          color: selectedColor,
          user: this.upnValue,
        };

        fetch(
          `${MatrixDataviewHtmlFormatter.config.commentApiUrl}/${commentId}`,
          {
            method: "PUT",
            mode: "cors",
            headers: { "Content-Type": "application/json" },
            body: JSON.stringify(updatedComment),
          }
        )
          .then(() => this.refreshCommentsAndRedraw())
          .catch((err) => console.error("Failed to update comment", err));
      } else {
        const selectedColor =
          (
            popup.querySelector(
              'input[name="commentColor"]:checked'
            ) as HTMLInputElement
          )?.value ?? "default";

        const commentData = {
          id: this.generateUUID(),
          user: this.upnValue,
          comment: commentText,
          column: columnHeaderText,
          filterId: filterId,
          createdAt: new Date().toISOString(),
          modifiedAt: new Date().toISOString(),
          color: selectedColor,
        };

        fetch(MatrixDataviewHtmlFormatter.config.commentApiUrl, {
          method: "POST",
          mode: "cors",
          headers: { "Content-Type": "application/json" },
          body: JSON.stringify(commentData),
        })
          .then(() => this.refreshCommentsAndRedraw())
          .catch((err) => console.error("Failed to submit comment", err));
      }

      this.target.removeChild(popup);
      textarea.dataset.commentId = "";
      textarea.dataset.editIndex = "";
    };

    // Append popup to the visual
    this.target.appendChild(popup);
  }

  public enumerateObjectInstances(
    options: EnumerateVisualObjectInstancesOptions
  ): VisualObjectInstance[] | VisualObjectInstanceEnumerationObject {
    const enumeration = new ObjectEnumerationBuilder();

    // Visuals are initialized with an empty data view before queries are run, therefore we need to make sure that
    // we are resilient here when we do not have data view.
    if (this.dataView) {
      let objects = null;
      if (this.dataView && this.dataView.metadata) {
        objects = this.dataView.metadata.objects;
      }

      switch (options.objectName) {
        case "general":
          break;
        case SubtotalProperties.ObjectSubTotals:
          this.enumerateSubTotalsOptions(enumeration, objects);
          break;
        case "visualSettings":
          enumeration.pushInstance({
            objectName: "visualSettings",
            properties: {
              redBarThreshold: objects?.visualSettings?.redBarThreshold ?? 0,
              redBarColumnsIndex:
                objects?.visualSettings?.redBarColumnsIndex ?? "",
              commentApiUrl: objects?.visualSettings?.commentApiUrl ?? "",
            },
            selector: null,
          });
          break;
        default:
          break;
      }
    }

    return enumeration.complete();
  }

  public enumerateSubTotalsOptions(
    enumeration,
    objects: DataViewObjects
  ): void {
    let instance = this.createVisualObjectInstance(
      SubtotalProperties.ObjectSubTotals
    );
    const rowSubtotalsEnabled: boolean = Visual.setInstanceProperty(
      objects,
      SubtotalProperties.rowSubtotals,
      instance
    );
    const columnSubtotalsEnabled: boolean = Visual.setInstanceProperty(
      objects,
      SubtotalProperties.columnSubtotals,
      instance
    );
    enumeration.pushInstance(instance);

    if (rowSubtotalsEnabled) {
      // Per row level
      instance = this.createVisualObjectInstance(
        SubtotalProperties.ObjectSubTotals
      );
      const perLevel = Visual.setInstanceProperty(
        objects,
        SubtotalProperties.rowSubtotalsPerLevel,
        instance
      );
      enumeration.pushInstance(instance, /* mergeInstances */ false);
      if (perLevel) {
        this.enumeratePerLevelSubtotals(
          enumeration,
          this.dataView.matrix.rows.levels
        );
      }
    }
    if (columnSubtotalsEnabled) {
      // Per column level
      instance = this.createVisualObjectInstance(
        SubtotalProperties.ObjectSubTotals
      );
      const perLevel = Visual.setInstanceProperty(
        objects,
        SubtotalProperties.columnSubtotalsPerLevel,
        instance
      );
      enumeration.pushInstance(instance, /* mergeInstances */ false);
      if (perLevel) {
        this.enumeratePerLevelSubtotals(
          enumeration,
          this.dataView.matrix.columns.levels
        );
      }
    }
  }

  private enumeratePerLevelSubtotals(
    enumeration,
    hierarchyLevels: DataViewHierarchyLevel[]
  ) {
    for (const level of hierarchyLevels) {
      for (const source of level.sources) {
        if (!source.isMeasure) {
          const instance = this.createVisualObjectInstance(
            SubtotalProperties.ObjectSubTotals,
            { metadata: source.queryName },
            source.displayName
          );
          Visual.setInstanceProperty(
            source.objects,
            SubtotalProperties.levelSubtotalEnabled,
            instance
          );
          enumeration.pushInstance(instance, /* mergeInstances */ false);
        }
      }
    }
  }

  private createVisualObjectInstance(
    objectName: string,
    selector: Selector = null,
    displayName?: string
  ): VisualObjectInstance {
    const instance: VisualObjectInstance = {
      selector: selector,
      objectName: objectName,
      properties: {},
    };

    if (displayName != null) instance.displayName = displayName;

    return instance;
  }

  private static getPropertyValue<T>(
    objects: DataViewObjects,
    dataViewObjectPropertyReference: DataViewObjectPropertyReference<T>
  ): T {
    let object;
    if (objects) {
      object =
        objects[dataViewObjectPropertyReference.propertyIdentifier.objectName];
    }
    return Visual.getValue(
      object,
      dataViewObjectPropertyReference.propertyIdentifier.propertyName,
      dataViewObjectPropertyReference.defaultValue
    );
  }

  private static setInstanceProperty<T>(
    objects: DataViewObjects,
    dataViewObjectPropertyReference: DataViewObjectPropertyReference<T>,
    instance: VisualObjectInstance
  ): T {
    const value = this.getPropertyValue(
      objects,
      dataViewObjectPropertyReference
    );
    if (instance && instance.properties) {
      instance.properties[
        dataViewObjectPropertyReference.propertyIdentifier.propertyName
      ] = value;
    }
    return value;
  }

  private static getValue<T>(
    object: DataViewObject,
    propertyName: string,
    defaultValue?: T,
    instanceId?: string
  ): T {
    if (!object) return defaultValue;

    if (instanceId) {
      const instances = object.$instances;
      if (!instances) return defaultValue;

      const instance = instances[instanceId];
      if (!instance) return defaultValue;

      object = instance;
    }

    const propertyValue = <T>object[propertyName];
    if (propertyValue === undefined) return defaultValue;

    return propertyValue;
  }
}
