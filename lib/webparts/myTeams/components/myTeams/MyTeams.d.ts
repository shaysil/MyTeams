/// <reference types="react" />
import * as React from 'react';
import { IMyTeamsProps, IMyTeamsState } from '.';
export declare class MyTeams extends React.Component<IMyTeamsProps, IMyTeamsState> {
    private _myTeams;
    constructor(props: IMyTeamsProps);
    componentDidMount(): Promise<void>;
    componentDidUpdate(prevProps: IMyTeamsProps): Promise<void>;
    private _load;
    render(): React.ReactElement<IMyTeamsProps>;
    private _onRenderCell;
    private _openChannel;
    private _getTeams;
    private _getTeamChannels;
}
