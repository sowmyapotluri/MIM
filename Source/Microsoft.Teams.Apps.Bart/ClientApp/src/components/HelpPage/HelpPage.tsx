import * as React from 'react';
import { Header, Menu } from '@stardust-ui/react';
import './HelpPage.scss';
import { RouteComponentProps } from 'react-router-dom'
import * as microsoftTeams from "@microsoft/teams-js";

export interface IHelpPageState {
    tabIndex: number;
}
var loginHint = "";
export default class HelpPage extends React.Component<RouteComponentProps, IHelpPageState> {
    constructor(props: RouteComponentProps) {
        super(props);
        this.state = {
            tabIndex: 0
        };
    }

    public componentDidMount = () => {
        microsoftTeams.initialize();
        microsoftTeams.getContext(context => {
           loginHint = context.upn ? context.upn : "";
           let teamID = context.teamId ? context.teamId : "";
           let teamName = context.teamName ? context.teamName : "";
           //let groupID = context.groupId ? context.groupId : "";
            try {
                console.log("teamID  " + teamID);
                console.log("teamName  " + teamName);
                console.log("loginHint  " + loginHint);
            }
            catch (error) {
                console.log("error in get channels");
            }
        });

    };

    public render(): JSX.Element {
        return (
            <div className="mainComponentHelpPage">
                <div className="header">
                    <Header key="dl"
                        color="brand"
                        content="Let's Get started ."
                    />
                    {loginHint}
                </div>
               
            </div>
        );
    }
}