export interface IServiceProps {
    context: any;
    Apilink: string;
    subscriptionId: string;
    OcpApimKey: string;
    Subject: string;
}
export interface IUserProfile {
    displayName: string;
    jobTitle: string;
    department: string;
    employeeId: string;
}
export interface IServiceRequestFormData {
    requestedBy?: string;
    requestedFor: string;
    requestedFor_Title: string;
    requestedFor_key: string;
    serviceName: string;
    serviceName_key: string;
    officeLocation: string;
    PhoneNumber: string;
    AccessthroughVPN: string;
    description: string;
    files?: any;
    AccessDurationFrom: Date;
    AccessDurationTo: Date;
}
//# sourceMappingURL=IServiceProps.d.ts.map