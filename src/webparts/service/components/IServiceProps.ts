export interface IServiceProps {
  context: any;
  Apilink: string;
  subscriptionId: string;
  OcpApimKey: string;
  Subject: string;
   isredirect:boolean;

    attachmentApilink: string;
  UserRecIdApilink: string;
  Category: string;
}
export interface IUserProfile {
  displayName: string;
  jobTitle: string;
  department: string;
  employeeId: string;
}
export interface IServiceRequestFormData {
  requestedBy?: string,
  requestedFor: string,
  requestedFor_Title: string,
  requestedFor_key: string,
  serviceName: string,
  serviceName_key: string,
  officeLocation: string,
  PhoneNumber: string,
  AccessthroughVPN:string;
  description: string;
  files?: any;
  AccessDurationFrom: Date,
  AccessDurationFroms: string,
    AccessDurationTos: string,
    AccessDurationTo: Date,
}