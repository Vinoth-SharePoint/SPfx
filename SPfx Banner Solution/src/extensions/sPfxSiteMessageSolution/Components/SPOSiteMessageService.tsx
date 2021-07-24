import * as React from 'react';
import { Link, MessageBar, MessageBarType } from '@microsoft/office-ui-fabric-react-bundle';
import { sp, Web } from "@pnp/sp/presets/all";
import { useEffect, useState } from 'react';
import * as strings from 'SPfxSiteMessageSolutionApplicationCustomizerStrings';
import { QUALIFIED_NAME } from '../SPfxSiteMessageSolutionApplicationCustomizer';

interface IBannerMessageItem {
    BannerMessage: string;
    Importance: string;
}

export interface IBannerMessageItemProps {
    listName: string;

}

export default function RenderSiteMessageService(props: IBannerMessageItemProps) {

    let [BannerMessages, SetBannerMessages] = useState<IBannerMessageItem[]>([]);
    useEffect(() => {
        // Use PnP JS to query SharePoint
        const now: string = new Date().toISOString();
        sp.web
            .lists.getByTitle(props.listName)
            .items
            .filter(`(StartDate le datetime'${now}' or StartDate eq null) and (EndDate ge datetime'${now}' or EndDate eq null)`)
            .select("BannerMessage", "StartDate", "EndDate", "Importance")
            .get<IBannerMessageItem[]>()
            .then(SetBannerMessages);

    }, [props.listName]);


    const BannerMessagesElements = BannerMessages
        .map((BannerMessageItem: { Importance: any; BannerMessage: any; }) => {

            let businessImpact = BannerMessageItem.Importance;
            switch (businessImpact) {
                case "High":
                    return (
                        <MessageBar
                            messageBarType={MessageBarType.error}
                            isMultiline={false}>
                            <strong>{BannerMessageItem.BannerMessage}</strong>&nbsp;
                        </MessageBar>
                    );
                    break;
                case "Medium":
                    return (
                        <MessageBar
                            messageBarType={MessageBarType.severeWarning}
                            isMultiline={false}>
                           <strong>{BannerMessageItem.BannerMessage}</strong>&nbsp;
                        </MessageBar>
                    );
                    break;
                case "Low":
                    return (
                        <MessageBar
                            messageBarType={MessageBarType.warning}
                            isMultiline={false}>
                            <strong>{BannerMessageItem.BannerMessage}</strong>&nbsp;
                        </MessageBar>
                    );
                    break;
            }

        }

        );

    return <div>{BannerMessagesElements}</div>;

}