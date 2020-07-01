import * as React from 'react';
import * as microsoftTeams from "@microsoft/teams-js";
import { initializeIcons } from 'office-ui-fabric-react/lib/Icons';
import { Input, Loader, Button, Flex, FlexItem, Text, Icon, Dropdown as FluentDropdown, DropdownProps, Checkbox, Layout, Divider, Segment, Grid, Header, TextArea, CheckboxProps, DropdownItemProps, ShorthandValue, ComponentEventHandler } from '@fluentui/react';
import { DatePicker, DayOfWeek, IDatePickerStrings } from 'office-ui-fabric-react/lib/DatePicker';
import { mergeStyleSets } from 'office-ui-fabric-react/lib/Styling';
import { Dropdown, IDropdownOption } from 'office-ui-fabric-react/lib/Dropdown';
import "./searchDetails.scss";
import "./bootstrap-grid.css";
import { isNullOrUndefined } from 'util';

const todayDate: Date = new Date();

const controlClass = mergeStyleSets({
    control: {
        margin: '0 0 15px 0',
        maxWidth: '300px'
    }
});

const DayPickerStrings: IDatePickerStrings = {
    months: ['January', 'February', 'March', 'April', 'May', 'June', 'July', 'August', 'September', 'October', 'November', 'December'],
    shortMonths: ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'],
    days: ['Sunday', 'Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday'],
    shortDays: ['S', 'M', 'T', 'W', 'T', 'F', 'S'],

    goToToday: '',
    prevMonthAriaLabel: 'Go to previous month',
    nextMonthAriaLabel: 'Go to next month',
    prevYearAriaLabel: 'Go to previous year',
    nextYearAriaLabel: 'Go to next year',
    closeButtonAriaLabel: 'Close date picker'
};

interface ISearchDetailsState {
    startDate: Date,
    endDate: Date,
    firstDayOfWeek: DayOfWeek,
    isToggled: boolean,
    duration: string,
    defaultSelectedIndexStartTime: string,
    defaultSelectedIndexEndTime: string,
    isStartTimeValid: boolean,
    isEndTimeValid: boolean,
    createdDate: Date,
    loader: boolean,
    page: number,
    allTimeZonesAvailable: [],
    allLocationsAvailable: [],
    audioCF: boolean,
    videoCF: boolean,
    presentation: boolean,
    wifi: boolean,
    allAvailableRooms: IRoom[],
    meetingTitle: string,
    meetingDescription: string,
    isSearchFieldsFilled: boolean,
    selectedTimezone: string,
    selectedLocations: string[],
    minValue: number,
    maxValue: number,
    selectAll: Boolean
};

interface IRoom {
    BuildingEmail: string
    BuildingName: string
    Location: string
    PartitionKey: string
    RoomEmail: string
    RoomName: string
    RowKey: string
    Status: string
    Timestamp: string
    UserAdObjectId: string
    label: string
    sublabel: string
    value: string
}

export default class SearchDetails extends React.Component<any, ISearchDetailsState> {

    token?: string | null = null;

    constructor(props: any) {
        super(props);
        initializeIcons();
        let startDate: Date = new Date();
        startDate.setHours(8, 30, 0, 0);
        let endDate: Date = new Date();
        endDate.setHours(9, 0, 0, 0);

        this.state = {
            startDate: startDate,
            endDate: endDate,
            firstDayOfWeek: DayOfWeek.Sunday,
            isToggled: false,
            duration: "",
            defaultSelectedIndexStartTime: "08:30 AM",
            defaultSelectedIndexEndTime: "09:00 AM",
            isStartTimeValid: true,
            isEndTimeValid: true,
            createdDate: todayDate,
            loader: true,
            page: 0,
            allTimeZonesAvailable: [],
            allLocationsAvailable: [],
            audioCF: false,
            videoCF: false,
            presentation: false,
            wifi: false,
            allAvailableRooms: [],
            meetingTitle: "",
            meetingDescription: "",
            isSearchFieldsFilled: false,
            selectedTimezone: "",
            selectedLocations: [],
            minValue: 0,
            maxValue: 0,
            selectAll: false
        };
        // let search = window.location.search;
        // let params = new URLSearchParams(search);
        // this.token = params.get("token");
    };



    public componentDidMount = () => {
        this.token = this.props.token;
        this.getSupportedTimeZones();
    }

    getSupportedTimeZones = async () => {
        // this.setState({ timeZonesLoading: true });
        let request = new Request("/api/MeetingApi/GetSupportedTimeZonesAsync", {
            headers: new Headers({
                "Authorization": "Bearer " + this.token
            })
        });

        const supportedTimezoneResponse = await fetch(request);
        // this.setState({ timeZonesLoading: false });

        if (supportedTimezoneResponse.status === 401) {
            const response = await supportedTimezoneResponse.json();
            if (response) {
                // this.setState({
                //     errorResponseDetail: {
                //         errorMessage: response.message,
                //         statusCode: response.code,
                //     }
                // })
            }

            // this.setState({ authorized: false });
            // this.appInsights.trackEvent({ name: `Unauthorized` }, { User: this.userObjectId });
        }
        else if (supportedTimezoneResponse.status === 200) {
            const response = await supportedTimezoneResponse.json();
            if (response !== null) {
                console.log("Response", response)
                this.setState({ allTimeZonesAvailable: response, loader: false });
                // let self = this;
                // let tzResult = self.state.supportedTimeZones.find(function (tz) { return tz === self.userTimeZone });
                // if (tzResult) {
                // this.setState({ selectedTimeZone: self.userTimeZone });
                // this.saveUserTimeZone(self.userTimeZone);
                // }
                // else {
                // this.setMessage(this.state.resourceStrings.TimezoneNotSupported, Constants.ErrorMessageRedColor, false);
                // }
            }
        }
    }

    private formatTime = (date: Date) => {
        return new Intl.DateTimeFormat('en-US', { hour: '2-digit', minute: '2-digit', hour12: true }).format(new Date(date)).toString().toUpperCase()
    };

    private datetime = (date: Date, time: string) => {
        let timeFormat = this.convertTime12to24(time).split(':');
        return new Date(date.getFullYear(), date.getMonth(), date.getDate(), parseInt(timeFormat[0]), parseInt(timeFormat[1]))
    }

    private convertTime12to24 = (time12h: any) => {
        const isPM = time12h.indexOf('PM') !== -1;
        let [hours, minutes] = time12h.replace(isPM ? 'PM' : 'AM', '').split(':');

        if (isPM) {
            hours = parseInt(hours, 10) + 12;
            hours = hours === 24 ? 12 : hours;
        } else {
            hours = parseInt(hours, 10);
            hours = hours === 12 ? 0 : hours;
            if (String(hours).length === 1) hours = '0' + hours;
        }

        const time = [hours, minutes].join(':');

        return time;
    }

    private _onFormatDate = (date: Date | null | undefined): string => {
        const shortMonths = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'];
        if (date != null) {
            return shortMonths[date.getMonth()] + ' ' + ('0' + date.getDate()).slice(-2) + ', ' + (date.getFullYear());
        }
        return "";
    };

    private _onSelectStartDate = (date: Date | null | undefined): void => {
        if (date != null)
            this.setState({ startDate: date });
    };

    private _onSelectEndDate = (date: Date | null | undefined): void => {
        if (date != null)
            this.setState({ endDate: date });
    };

    private setStartTime = (e: React.SyntheticEvent<HTMLElement>, option?: IDropdownOption, index?: number) => {
        console.log(option!.text);
        if (option != null) {
            let date = this.datetime(this.state.startDate, option!.text);
            console.log(date);
            this.setState({
                defaultSelectedIndexStartTime: option!.text,
                startDate: date,
            },
                () => {
                    if (this.state.startDate > this.state.endDate) {
                        this.setState({
                            isStartTimeValid: false,
                            isEndTimeValid: false,
                        })
                    }
                    else {
                        this.setState({
                            isStartTimeValid: true,
                            isEndTimeValid: true,
                        })
                    }
                    this.duration();
                });
        }
    }

    private setEndTime = (e: React.SyntheticEvent<HTMLElement>, option?: IDropdownOption, index?: number) => {
        if (option != null) {
            let date = this.datetime(this.state.endDate, option!.text);
            this.setState({
                defaultSelectedIndexEndTime: option!.text,
                endDate: date
            },
                () => {
                    if (this.state.endDate > this.state.startDate) {
                        this.setState({
                            isEndTimeValid: true,
                            isStartTimeValid: true,
                        })
                    }
                    else {
                        this.setState({
                            isEndTimeValid: false,
                            isStartTimeValid: false,
                        })
                    }
                    this.duration();
                });
        }
    }

    private duration = () => {
        let startDate = this.state.startDate.getTime();
        let endDate = this.state.endDate.getTime();
        const diffInMinutes = Math.floor(Math.abs(startDate - endDate) / 60000);
        let diff = diffInMinutes < 60 ? diffInMinutes + "m" : (diffInMinutes < 60 * 24 ? (diffInMinutes / 60).toFixed(1) + "h" : (diffInMinutes / (60 * 24)).toFixed(1) + "d")
        this.setState({ duration: diff });
    }

    private onStartTimeChange = (e: React.SyntheticEvent<HTMLElement, Event>) => {
        let startTime = (e.target as HTMLInputElement).innerText;
        let regExp = /^(1[0-2]|0?[1-9]):[0-5][0-9] (AM|PM)$/i;
        let extractedTime = startTime.toLowerCase().split("m")[0].toLowerCase() + "m";
        if (regExp.test(extractedTime)) {
            let date = this.datetime(this.state.startDate, startTime);
            this.setState({
                isStartTimeValid: true,
                startDate: date
            },
                () => {
                    if (this.state.startDate > this.state.endDate) {
                        this.setState({
                            isStartTimeValid: false,
                            isEndTimeValid: false,
                        })
                    }
                    else {
                        this.setState({
                            isStartTimeValid: true,
                            isEndTimeValid: true,
                        })
                    }
                    this.duration();
                })

        }
        else {
            this.setState({ isStartTimeValid: false })
        }
    }

    private onEndTimeChange = (e: React.SyntheticEvent<HTMLElement, Event>) => {
        let endTime = (e.target as HTMLInputElement).innerText;
        let regExp = /^(1[0-2]|0?[1-9]):[0-5][0-9] (AM|PM)$/i;
        let extractedTime = endTime.toLowerCase().split("m")[0].toLowerCase() + "m";
        if (regExp.test(extractedTime)) {
            let date = this.datetime(this.state.startDate, endTime);
            this.setState({
                isEndTimeValid: true,
                endDate: date
            },
                () => {
                    if (this.state.endDate > this.state.startDate) {
                        this.setState({
                            isEndTimeValid: true,
                            isStartTimeValid: true,
                        })
                    }
                    else {
                        this.setState({
                            isEndTimeValid: false,
                            isStartTimeValid: false,
                        })
                    }
                    this.duration();
                })
        }
        else {
            this.setState({ isEndTimeValid: false })
        }
    }

    private onCheckboxChanged = (e: React.SyntheticEvent<HTMLElement, Event>, v?: CheckboxProps) => {
        console.log("Event", e);
        console.log("CheckboxProps", v);
        let checked = v!.checked!;
        switch (v!.label!) {
            case "Video CF":
                this.setState({
                    videoCF: checked,
                }, () => {
                    this.fieldsValidors();
                });
                break;
            case "Audio Room":
                this.setState({
                    audioCF: checked,
                }, () => {
                    this.fieldsValidors();
                });
                break;
            case "Pictures available":
                this.setState({
                    wifi: checked
                }, () => {
                    this.fieldsValidors();
                });
                break;
            case "Presentation":
                this.setState({
                    presentation: checked,
                }, () => {
                    this.fieldsValidors();
                });
                break;
        }
    }

    private fieldsValidors = () => {
        console.log("loc", this.state.selectedLocations)
        console.log("loc2", this.state.selectedTimezone)
        if ((this.state.audioCF || this.state.videoCF || this.state.presentation ||
            this.state.wifi) && this.state.selectedLocations.length > 0 &&
            this.state.selectedTimezone !== "") {
            this.setState({
                isSearchFieldsFilled: true
            });
        }
        else {
            this.setState({
                isSearchFieldsFilled: false
            });
        }
    }

    private selectLocations = {
        // onAdd: ((item: ShorthandValue<DropdownItemProps>) => console.log("Item added", item)),
        // onRemove: ((item: ShorthandValue<DropdownItemProps>) => console.log("Item removed", item)),
        onAdd: (item: ShorthandValue<DropdownItemProps>) => {
            console.log("Item added", item);
            var listOfLoctions = this.state.selectedLocations;
            listOfLoctions.push(item!.toString()!);
            this.setState({
                selectedLocations: listOfLoctions
            }, () => {
                this.fieldsValidors();
            });
            return `${item} has been selected.`
        },
        onRemove: (item: any) => {
            console.log("Item removed", item);
            var listOfLoctions = this.state.selectedLocations.splice(this.state.selectedLocations.indexOf(item!.toString()), 1);
            this.setState({
                selectedLocations: listOfLoctions
            }, () => {
                this.fieldsValidors();
            });
            return `${item} has been removed.`
        }
    }

    private timezoneChange = (e: React.SyntheticEvent<HTMLElement, Event>, v?: DropdownProps) => {
        console.log("V", v!.value)
        this.setState({
            selectedTimezone: String(v!.value!)
        });
    }

    private minValueChanged = (e: React.SyntheticEvent<HTMLElement, Event>) => {
        let min = Number((e.target as HTMLInputElement).value);
        this.setState({
            minValue: min
        });
    }

    private maxValueChanged = (e: React.SyntheticEvent<HTMLElement, Event>) => {
        let max = Number((e.target as HTMLInputElement).value);
        this.setState({
            maxValue: max
        });
    }

    private onSelectAllCheckboxChanged = (e: React.SyntheticEvent<HTMLElement, Event>, v?: CheckboxProps) => {
        console.log("Event", e);
        console.log("CheckboxProps", v);
        let checked = v!.checked!;
        this.setState({
            selectAll: checked,
            audioCF: checked,
            videoCF: checked,
            presentation: checked,
            wifi: checked
        }, () => {
            this.fieldsValidors();
        });
    }

    private searchRooms = () => {
        this.setState({
            loader: true
        });
        this.searchRoomByFeatures();
    }

    /**
    * Filter rooms as per user input.
    * @param inputValue Input string.
    */
    private searchRoomByFeatures = async () => {
        let self = this;
        let rooms = { Location: this.state.selectedLocations.join(','), Minimum: this.state.minValue, Maximum: this.state.maxValue, Time: "2020-06-03 11:00:00", Duration: 2, Timezone: "US/Central" };
        await fetch("/api/MeetingApi/SearchRoomByFeaturesAsync", {
            method: "POST",
            headers: {
                "Content-Type": "application/json",
                "Authorization": "Bearer " + this.token
            },
            body: JSON.stringify(rooms)
        }).then(async (res) => {
            if (res.status === 401) {
                const response = await res.json();
                if (response) {
                    // this.setState({
                    //     errorResponseDetail: {
                    //         errorMessage: response.message,
                    //         statusCode: response.code,
                    //     }
                    // })
                }

                // this.setState({ authorized: false });
                // this.appInsights.trackTrace({ message: `User ${this.userObjectIdentifier} is unauthorized!`, severityLevel: SeverityLevel.Warning });
                return response;
            }
            else if (res.status === 200) {
                let response = await res.json();
                let rooms: IRoom[] = [];
                for (let j: number = 0; j < response.length; j++) {
                    rooms.push(
                        {
                            BuildingEmail: response[j].BuildingEmail,
                            BuildingName: response[j].BuildingName,
                            Location: response[j].Location,
                            PartitionKey: response[j].PartitionKey,
                            RoomEmail: response[j].RoomEmail,
                            RoomName: response[j].RoomName,
                            RowKey: response[j].RowKey,
                            Status: response[j].Status,
                            Timestamp: response[j].Timestamp,
                            UserAdObjectId: response[j].UserAdObjectId,
                            label: response[j].label,
                            sublabel: response[j].sublabel,
                            value: response[j].value,
                        }
                    );
                }
                this.setState({
                    allAvailableRooms: rooms
                }, () => {
                    this.setState({
                        page: 1,
                        loader: false
                    });
                });
            }
            else {
                // this.setMessage(this.state.resourceStrings.ExceptionResponse, Constants.ErrorMessageRedColor, false);
                // this.appInsights.trackTrace({ message: `'SearchRoomAsync' - Request failed:${res.status}`, severityLevel: SeverityLevel.Warning });
            }

        });
    }

    private getLocations = (e: React.SyntheticEvent<HTMLElement, Event>, data?: DropdownProps) => {
        var searchQuery = data!.searchQuery!
        this.filterRooms(searchQuery).then((res: any) => {
            this.setState({
                allLocationsAvailable: res
            });
        });
    }

    /**
    * Filter rooms as per user input.
    * @param inputValue Input string.
    */
    private filterRooms = async (inputValue: string) => {
        let self = this;

        if (inputValue) {
            let rooms = { Query: inputValue };
            const res = await fetch("/api/MeetingApi/SearchLocationAsync", {
                method: "POST",
                headers: {
                    "Content-Type": "application/json",
                    "Authorization": "Bearer " + this.token
                },
                body: JSON.stringify(rooms)
            });

            if (res.status === 401) {
                const response = await res.json();
                if (response) {
                    // this.setState({
                    //     errorResponseDetail: {
                    //         errorMessage: response.message,
                    //         statusCode: response.code,
                    //     }
                    // })
                }

                // this.setState({ authorized: false });
                // this.appInsights.trackTrace({ message: `User ${this.userObjectIdentifier} is unauthorized!`, severityLevel: SeverityLevel.Warning });
                return response;
            }
            else if (res.status === 200) {
                let response = await res.json();
                console.log("Lsit", response)
                return response;
            }
            else {
                // this.setMessage(this.state.resourceStrings.ExceptionResponse, Constants.ErrorMessageRedColor, false);
                // this.appInsights.trackTrace({ message: `'SearchRoomAsync' - Request failed:${res.status}`, severityLevel: SeverityLevel.Warning });
            }
        }
    }

    public render(): JSX.Element {
        let items = [] //Populate grid items
        for (let i: number = 0; i < this.state.allAvailableRooms.length; i++) {
            items.push(<Segment >
                <Flex gap="gap.small">
                    <FlexItem grow>
                        <Checkbox key={this.state.allAvailableRooms[i].RowKey} id={this.state.allAvailableRooms[i].RowKey} label={this.state.allAvailableRooms[i].RoomName}/> 
                    </FlexItem>
                </Flex>
            </Segment>)
            items.push(<Segment content={this.state.allAvailableRooms[i].Location} ></Segment>)
            items.push(<Segment content={this.state.allAvailableRooms[i].UserAdObjectId} ></Segment>)
        }

        let minInterval = 30; //minutes interval
        const times = []; // time array
        let startTime = 0; // start time
        let ap = ['AM', 'PM']; // AM-PM

        //loop to increment the time and push results in array
        for (let i = 0; startTime < 24 * 60; i++) {
            let hh = Math.floor(startTime / 60); // getting hours of day in 0-24 format
            let mm = (startTime % 60); // getting minutes of the hour in 0-55 format
            times[i] = ("0" + (hh % 12)).slice(-2) + ':' + ("0" + mm).slice(-2) + ' ' + ap[Math.floor(hh / 12)]; // pushing data in array in [00:00 - 12:00 AM/PM format]
            startTime = startTime + minInterval;
        }
        const timeData: IDropdownOption[] = [];

        for (let i = 0; i < times.length; i++) {
            timeData.push({ key: times[i], text: times[i] });
        }

        let timezoneContainer = (
            <div>
                <Flex>
                    <Text content="Timezone" />
                </Flex>
                <Flex gap="gap.smaller">
                    <FluentDropdown
                        items={this.state.allTimeZonesAvailable}
                        placeholder="Select your timezone"
                        onSelectedChange={this.timezoneChange}
                    />
                </Flex>
            </div>
        );

        let fromContainer = (
            <div>
                <Flex>
                    <Text content="From" />
                </Flex>
                <Flex gap="gap.smaller">
                    <DatePicker
                        className={controlClass.control}
                        strings={DayPickerStrings}
                        showWeekNumbers={false}
                        firstWeekOfYear={1}
                        showMonthPickerAsOverlay={true}
                        placeholder="Select a date..."
                        ariaLabel="Select a date"
                        formatDate={this._onFormatDate}
                        minDate={todayDate}
                        onSelectDate={this._onSelectStartDate}
                        value={this.state.startDate}
                    />
                    <Dropdown
                        contentEditable
                        onInput={this.onStartTimeChange}
                        defaultSelectedKey={this.state.defaultSelectedIndexStartTime}
                        options={timeData}
                        placeholder="select start time"
                        disabled={this.state.isToggled}
                        onChange={this.setStartTime}
                        errorMessage={!this.state.isStartTimeValid ? 'Please provide valid time' : ''}
                    />
                </Flex>
            </div>
        );

        let toContainer = (
            <div>
                <Flex>
                    <Text content="To" />
                </Flex>
                <Flex gap="gap.smaller">
                    <DatePicker
                        className={controlClass.control}
                        strings={DayPickerStrings}
                        showWeekNumbers={false}
                        firstWeekOfYear={1}
                        showMonthPickerAsOverlay={true}
                        placeholder="Select a date..."
                        ariaLabel="Select a date"
                        formatDate={this._onFormatDate}
                        minDate={this.state.startDate}
                        onSelectDate={this._onSelectEndDate}
                        value={this.state.endDate}
                    />
                    <Dropdown
                        contentEditable
                        onInput={this.onEndTimeChange}
                        defaultSelectedKey={this.state.defaultSelectedIndexEndTime}
                        options={timeData}
                        placeholder="select end time"
                        onChange={this.setEndTime}
                        disabled={this.state.isToggled}
                        errorMessage={!this.state.isEndTimeValid ? 'Please provide valid time' : ''}
                    />
                </Flex>
            </div>
        );

        if (this.state.loader) {
            return (
                <div className="emptyContent">
                    <Loader />
                </div>
            );
        }
        else {
            if (this.state.page === 0) {
                return (
                    <div className="taskModule">
                        <div className="formContainer">
                            <Flex gap="gap.medium" column>
                                <Flex>
                                    <Layout gap="2rem" start={timezoneContainer} main={fromContainer} end={toContainer} />;
                                </Flex>
                                <Flex gap="gap.large">
                                    <Flex column gap="gap.medium">
                                        <Text content="Location" />
                                        <FlexItem>
                                            <FluentDropdown
                                                search
                                                multiple
                                                loading={this.state.allLocationsAvailable === []}
                                                onSearchQueryChange={this.getLocations}
                                                items={this.state.allLocationsAvailable}
                                                placeholder="Start typing a name"
                                                getA11ySelectionMessage={this.selectLocations}
                                                noResultsMessage="We couldn't find any matches."
                                            />
                                        </FlexItem>
                                        <Text content="Capacity" />
                                        <Flex gap="gap.smaller">
                                            <Flex gap="gap.smaller">
                                                <Text content="Min" />
                                                <FlexItem>
                                                    <Input type="number" min={0} onChange={this.minValueChanged} />
                                                </FlexItem>
                                            </Flex>
                                            <Flex gap="gap.smaller">
                                                <Text content="Max" />
                                                <FlexItem>
                                                    <Input type="number" min={this.state.minValue === 0 ? 0 : this.state.minValue} onChange={this.maxValueChanged} />
                                                </FlexItem>
                                            </Flex>
                                        </Flex>
                                    </Flex>
                                    <FlexItem>
                                        <Flex gap="gap.medium" column>
                                            <Flex gap="gap.large">
                                                <Text content="Capabilities" />
                                                <FlexItem push>
                                                    <Checkbox label="Select all" onClick={this.onSelectAllCheckboxChanged} />
                                                </FlexItem>
                                            </Flex>
                                            <Flex gap="gap.large">
                                                <Flex column gap="gap.small">
                                                    <Checkbox label="Audio Room" onClick={this.onCheckboxChanged} checked={this.state.audioCF} />
                                                    <Checkbox label="Video CF" onClick={this.onCheckboxChanged} checked={this.state.videoCF} />
                                                </Flex>
                                                <FlexItem>
                                                    <div>
                                                        <Flex column gap="gap.small">
                                                            <Checkbox label="Pictures available" onClick={this.onCheckboxChanged} checked={this.state.wifi} />
                                                            <Checkbox label="Presentation" onClick={this.onCheckboxChanged} checked={this.state.presentation} />
                                                        </Flex>
                                                    </div>
                                                </FlexItem>
                                            </Flex>
                                        </Flex>
                                    </FlexItem>
                                </Flex>
                                {/* <Flex gap="gap.small">
                                    <FlexItem push>
                                    </FlexItem>
                                </Flex> */}
                            </Flex>
                            <div className="footerContainer">
                                <div className="buttonContainer">
                                    <Flex>
                                        {this.state.duration}
                                        <FlexItem push>
                                            <Button content="Search rooms" primary disabled={!this.state.isSearchFieldsFilled} onClick={this.searchRooms} />
                                        </FlexItem>
                                    </Flex>
                                </div>
                            </div>
                        </div>
                    </div>
                );

            }
            else if (this.state.page === 1) {
                return (
                    <div className="listComponent">
                        <div className="formContentSecondContainer GridBrdPD mr-top-sm" >
                            <Grid columns=".1fr 2.5fr 3fr 1.5fr">
                                <Segment color="brand" >
                                    <Flex gap="gap.small" className="GridTitle" color="brand" >
                                        <FlexItem grow>
                                            <Checkbox key="all" id="all" />
                                        </FlexItem>
                                    </Flex>
                                </Segment>
                                <Segment color="brand" >
                                    <Flex gap="gap.small" className="GridTitle">
                                        <FlexItem>
                                            <Text content="Room name" />
                                        </FlexItem>
                                    </Flex>
                                </Segment>
                                <Segment color="brand" >
                                    <Flex gap="gap.small" className="GridTitle">
                                        <FlexItem grow>
                                            <Text content="Location" />
                                        </FlexItem>
                                    </Flex>
                                </Segment>
                                <Segment color="brand" >
                                    <Flex gap="gap.small" className="GridTitle">
                                        <FlexItem grow>
                                            <Text content="Room ID" />
                                        </FlexItem>
                                    </Flex>
                                </Segment>
                                {items}
                            </Grid>
                        </div>
                    </div>
                );
            }
            else {
                return (
                    <div className="taskModule">
                        <Divider content="Next page" />
                        <div className="formContainer">
                            <Header
                                as="h3"
                                content="Meeting details"
                                description={{
                                    content: 'Almost done! Please add these details so that i can create a meeting for you.',
                                    as: 'p',
                                }} />
                            <div className="MeetingDetails">
                                <Flex column gap="gap.large">
                                    <Flex column>
                                        <Text content="Meeting title" />
                                        <Input fluid />
                                    </Flex>
                                    <Flex column>
                                        <Text content="Description (Optional)" />
                                        <TextArea />
                                    </Flex>
                                    <Flex column>
                                        <Text content="Location" />
                                        <TextArea disabled />
                                    </Flex>
                                </Flex>
                            </div>
                        </div>
                    </div>
                );

            }
        }
    }
}