//https://graph.microsoft.com/v1.0/me?$select=displayName,companyName,jobTitle,mobilePhone,mail
export interface ProfileCardData {
  title: string;
  displayName: string;
  profileImage: string;
  companyName: string;
  jobTitle: string;
  properties: ProfileCardDataProperty[]
}

export interface ProfileCardDataProperty {
  key: string;
  value: string;
}