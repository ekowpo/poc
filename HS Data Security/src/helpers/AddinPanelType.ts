import { buildUIReadPane } from "./../taskpane/Read";
import { writeDataToOfficeDocument } from "./../taskpane/taskpane";

export const ReadPane = "Readpane";
export const TaskPane = "Taskpane";

export function paneToDisplay(pane: string, result: object) {
  switch (pane) {
    case ReadPane:
      buildUIReadPane(result);
      break;
    case TaskPane:
      writeDataToOfficeDocument(result);
      break;
  }
}
