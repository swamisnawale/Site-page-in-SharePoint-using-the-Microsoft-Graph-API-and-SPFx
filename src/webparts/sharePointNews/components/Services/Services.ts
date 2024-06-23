import { WebPartContext } from "@microsoft/sp-webpart-base";
import { getSP } from "./pnpJs";

export const addItemInList = async (
  props: {
    context: WebPartContext;
    listName: string;
  },
  newItemBody: any
) => {
  const sp = await getSP(props.context);
  try {
    const items: any = await sp.web.lists
      .getByTitle(`${props.listName}`)
      .items.add(newItemBody);
    return items;
  } catch (error) {
    const errorJson = await error.response.json();
    throw new Error(`Failed to add item: ${errorJson.error.message}`);
  }
};
