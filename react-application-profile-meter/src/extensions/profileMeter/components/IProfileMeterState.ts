import IUser from "./IUser";

export default interface IProfileMeterState {
    score: number;
    showPanel: boolean;
    currentUser: IUser;
}