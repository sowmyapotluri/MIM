/*
    <copyright file="add-favorites.tsx" company="Microsoft Corporation">
    Copyright (c) Microsoft Corporation. All rights reserved.
    </copyright>
*/

import * as React from "react";
import "./theme.css";
import { Button, Loader, Divider, List, Icon, Text, Provider, themes, Dropdown, Flex, Menu } from '@fluentui/react';
import { components } from "react-select";
import AsyncSelect from "react-select/async";
import * as microsoftTeams from "@microsoft/teams-js";
import * as Constants from "./constants";
import { ApplicationInsights, SeverityLevel } from '@microsoft/applicationinsights-web';
import { ReactPlugin, withAITracking } from '@microsoft/applicationinsights-react-js';
import { createBrowserHistory } from "history";
import SearchDetails from './searchDetails';
import * as timezone from 'moment-timezone';
import moment from 'moment';
const browserHistory = createBrowserHistory({ basename: '' });
interface IAddFavoriteProps { }
let reactPlugin = new ReactPlugin();

/** User favorite room */
class FavoriteRoom {
    UserAdObjectId?: string | null = null;
    RoomEmail?: string | null = null;
    RoomName?: string | null = null;
    BuildingName?: string | null = null;
    BuildingEmail?: string | null = null;
}

/** State interface. */
interface IState {
    /**Favorite rooms list. */
    favoriteRooms: Array<FavoriteRoom>,
    /**Selected room object. */
    selectedRoom: any,
    /**Loading icon visibility. */
    loading: boolean,
    /**Add button disable/enable. */
    addDisable: boolean,
    /**Error message text. */
    message?: string | null,
    /**Error message visibility. */
    showMessage: boolean,
    /**Error message color. */
    messageColor?: string,
    /**Selected theme. */
    theme: any,
    /**Is user authorized. */
    authorized: boolean,
    /**Loading for favorite list. */
    loadingFavoriteList: boolean,
    /**Top five rooms to display in dropdown */
    topFiveRooms: Array<any>,
    /**Supported time zones for user */
    supportedTimeZones: Array<any>,
    /**Selected time zone for user */
    selectedTimeZone: any,
    /**Boolean indicating if time zones are loading in dropdown */
    timeZonesLoading: boolean,
    resourceStrings: any,
    resourceStringsLoaded: boolean,
    isRoomDeleted: boolean,
    errorResponseDetail: IErrorResponse,
    tabIndex: number
};

/**Server error response interface */
interface IErrorResponse {
    statusCode?: string,
    errorMessage?: string,
}

/** Component for managing user favorites. */
class FindRoom extends React.Component<IAddFavoriteProps, IState>
{
    /**Reply to activity Id. */
    replyTo?: string | null = null;
    /**Auth token. */
    token?: string | null = null;
    /**Component state. */
    state: IState;
    /** Theme color according to teams theme*/
    themeColor?: any = undefined;
    /** Theme styles according to teams theme*/
    themeStyle?: any = undefined;
    /** Instrumentation key for telemetry logging*/
    telemetry: any = undefined;
    appInsights: ApplicationInsights;
    userObjectId: any;
    userTimeZone: any = null;
    strings: any = {};

    /**
     * Contructor to initialize component.
     * @param props Props of component.
     */
    constructor(props: IAddFavoriteProps) {
        super(props);
        this.state = {
            favoriteRooms: [],
            selectedRoom: null,
            loading: false,
            addDisable: true,
            message: null,
            showMessage: false,
            messageColor: undefined,
            theme: null,
            authorized: true,
            loadingFavoriteList: false,
            topFiveRooms: [],
            supportedTimeZones: [],
            selectedTimeZone: null,
            timeZonesLoading: false,
            resourceStrings: {},
            resourceStringsLoaded: false,
            isRoomDeleted: false,
            errorResponseDetail: {
                errorMessage: undefined,
                statusCode: undefined,
            },
            tabIndex: 0
        };

        let search = window.location.search;
        let params = new URLSearchParams(search);
        this.telemetry = params.get("telemetry");
        this.token = params.get("token");
        this.appInsights = new ApplicationInsights({
            config: {
                instrumentationKey: this.telemetry,
                extensions: [reactPlugin],
                extensionConfig: {
                    [reactPlugin.identifier]: { history: browserHistory }
                }
            }
        });
        this.appInsights.loadAppInsights();
    }

    /** Called once component is mounted. */
    async componentDidMount() {
        this.setState({ loading: true });

        // Call the initialize API first
        microsoftTeams.initialize();

        // Check the initial theme user chose and respect it
        microsoftTeams.getContext((context) => {
            this.userObjectId = context.userObjectId;
            if (context && context.theme) {
                this.setState({ theme: context.theme });
                if (context.theme === Constants.DarkTheme) {
                    this.themeColor = Constants.DarkThemeColors;
                    this.themeStyle = Constants.DarkStyles;
                }
                else if (context.theme === Constants.DefaultTheme) {
                    this.themeColor = Constants.DefaultThemeColors;
                    this.themeStyle = Constants.DefaultStyles;
                }
                else {
                    this.themeColor = Constants.ContrastThemeColors;
                    this.themeStyle = Constants.ContrastStyles;
                }
            }
        });

        // Handle theme changes
        microsoftTeams.registerOnThemeChangeHandler((theme) => {
            this.setState({ theme: theme });
            if (theme === Constants.DarkTheme) {
                this.themeColor = Constants.DarkThemeColors;
                this.themeStyle = Constants.DarkStyles;
            }
            else if (theme === Constants.DefaultTheme) {
                this.themeColor = Constants.DefaultThemeColors;
                this.themeStyle = Constants.DefaultStyles;
            }
            else {
                this.themeColor = Constants.ContrastThemeColors;
                this.themeStyle = Constants.ContrastStyles;
            }
        });
    }

    private items: { key: string, content: string }[] = [
        {
            key: "findroom",
            content: "FIND ROOM"
        },
        {
            key: "fav",
            content: "FAVOURITES"
        }
    ];

    private tabClicked = (item: any, value: any) => {
        if (value.content === "FAVOURITES") {
            this.setState({
                tabIndex: 1
            });
        }
        else {
            this.setState({
                tabIndex: 0
            });
        }
    }


    /** Render function. */
    render() {
        let self = this;
        const checkAuthAndRender = function () {
            if (self.state.authorized) {
                // if (self.state.resourceStringsLoaded) {
                    return (
                        <Provider theme={self.state.theme === Constants.DefaultTheme ? themes.teams : self.state.theme === Constants.DarkTheme ? themes.teamsDark : themes.teamsHighContrast}>
                            <div className="containerdiv">
                                <Menu items={self.items} defaultActiveIndex={0} primary underlined onItemClick={self.tabClicked} />
                                {self.state.tabIndex === 0 ? <SearchDetails token={self.token}/> : <Text content="Under construction"/>}
                            </div>
                        </Provider>
                    );
                // }
                // else {
                //     return (<Loader />);
                // }
            }
            else {
                return (
                    <Provider theme={self.state.theme === Constants.DefaultTheme ? themes.teams : self.state.theme === Constants.DarkTheme ? themes.teamsDark : themes.teamsHighContrast}>
                        <div className="containerdiv">
                            <div className="containerdiv-unauthorized">
                                <Flex gap="gap.small" vAlign="center" hAlign="center">
                                    {/* {self.renderErrorMessage()} */}
                                </Flex>
                            </div>
                        </div>
                    </Provider>
                );
            }
        }

        return (checkAuthAndRender());
    }
}

export default withAITracking(reactPlugin, FindRoom);