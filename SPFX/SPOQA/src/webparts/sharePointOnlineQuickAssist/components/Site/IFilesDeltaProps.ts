import {MSGraphClient} from '@microsoft/sp-http';
export interface IFilesDeltaProps {
  description: string;
  msGraphClient:MSGraphClient;
}
