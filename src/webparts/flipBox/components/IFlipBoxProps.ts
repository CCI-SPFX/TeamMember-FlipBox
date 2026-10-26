
import { WebPartContext } from '@microsoft/sp-webpart-base';

export interface IFlipBoxProps {
  context: WebPartContext;
  listName: string;
  linkedInIconUrl: string;
  xIconUrl: string;
}