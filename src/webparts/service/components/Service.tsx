import * as React from "react";
import { useEffect, useState } from "react";
import { MSGraphClient } from "@microsoft/sp-http";
import styles from "./Service.module.scss";
import type { IServiceProps, IServiceRequestFormData } from "./IServiceProps";
import { escape } from "@microsoft/sp-lodash-subset";
import AlertModal from "../../../components/alertModal/AlertModal";
import { Web } from "@pnp/sp/webs";
import ServiceUIForm from "../../../components/ServiceUIForm";
interface IUserProfile {
  displayName: string;
  jobTitle: string;
  department: string;
  employeeId: string;
}
//const rootSiteURL = window.location.protocol + "//" + window.location.hostname + "/sites/MCIT-Internal-Services";
const getUserInitials = (displayName: string): string => {
  const names = displayName.trim().split(" ");
  const initials = names.map((name) => name.charAt(0).toUpperCase()).join("");
  return initials;
};
const generateGUID = (): string => {
  return "xxxxxxxx-xxxx-4xxx-yxxx-xxxxxxxxxxxx".replace(/[xy]/g, (c) => {
    const r = (Math.random() * 16) | 0;
    const v = c === "x" ? r : (r & 0x3) | 0x8;
    return v.toString(16);
  });
};
const generateUserTitle = async (
  userProfileAD: IUserProfile | null
): Promise<string> => {
  if (!userProfileAD || !userProfileAD.displayName) {
    throw new Error("User profile information is missing.");
  }
  const userInitials = getUserInitials(userProfileAD.displayName);
  const guid = generateGUID().substring(0, 8);
  const title = `MR-${userInitials}-${guid}`;
  console.log("Generated User Title:", title);
  return title;
};
const ServiceRequest: React.FC<IServiceProps> = (props) => {
  const [userProfileAD, setUserProfileAD] = useState<IUserProfile | null>(null);
  const [isLoadingUser, setIsLoadingUser] = useState<boolean>(true);
  const [showModal, setShowModal] = useState(false);
  const [modalHeading, setModalHeading] = useState("");
  const [modalMessage, setModalMessage] = useState("");
  const [alertsection, setAlertsection] = useState("");
  const [iconLoad, setIconLoad] = useState("");
  const handleShowModal = () => setShowModal(true);
  const handleCloseModal = (section: string) => {
    setShowModal(false);
  };

  useEffect(() => {
    (async () => {
      try {
        const client: MSGraphClient =
          await props.context.msGraphClientFactory.getClient("3");
        const userAD: any = await client
          .api("/me")
          .select(
            "displayName,jobTitle,department,employeeId,mail,onPremisesExtensionAttributes"
          )
          .get();

        const userProfile: IUserProfile = {
          displayName: userAD.displayName || "",
          jobTitle: userAD.jobTitle || "",
          department: userAD.department || "",
          employeeId:
            userAD?.onPremisesExtensionAttributes?.extensionAttribute15 || "",
        };

        setUserProfileAD(userProfile);
        setIsLoadingUser(false);
      } catch (error) {
        console.error("Error fetching user info:", error);
        setIsLoadingUser(false);
      }
    })();
  }, [props]);

  const showErrorModal = () => {
    setModalHeading("Warning");
    setModalMessage("Please fill Required fields");
    setAlertsection("rejected");
    setIconLoad("WarningSolid");
    handleShowModal();
  };

  const saveRequest = async (formData: IServiceRequestFormData) => {
    try {
      console.log(formData);
      const payload = {
        attachmentsToDelete: [],
        attachmentsToUpload: [],
        parameters: {
          "par-39A36FCD648F4578B1552F0D6AEF398C":formData.requestedBy,
          "par-FC7FDDC6F60C4AACAB58C193D7BE4231":formData.requestedFor_Title,
          "par-FC7FDDC6F60C4AACAB58C193D7BE4231-recId":formData.requestedFor_key,
          "par-7EE4B6BBC3DB474692536C0C34AC9DA2": formData.serviceName,
          "par-7EE4B6BBC3DB474692536C0C34AC9DA2-recId": formData.serviceName_key,
          "par-9466DA61B86B43CE94D84B2F8BF7C9C4": formData.officeLocation,
          "par-EF33DA0BF7D2473BA6A9458B69B9759D": formData.PhoneNumber,
          "par-555FDA87B1B44A93B4EACC62E32ECBCE": formData.AccessDurationFroms,
          "par-937670E4139346F59F31AEF4C7A26C81": formData.AccessDurationTos,
          "par-34988302FB614F19B14B68D9CDFD65B1": formData.AccessthroughVPN,
          "par-3C60AFC8919E40D8A499BE83BD729809": formData.description,
        },
        delayedFulfill: false,
        formName: "ServiceReq.ResponsiveAnalyst.DefaultLayout",
        saveReqState: false,
        serviceReqData: {
          Subject: `${props.Subject}`,
          Symptom: formData.description, // "It allows employees to make Mobile and International calls with standards features like Voicemail and Call Forwarding",
          Category: props.Category,// "Calling",
          CreatedBy:formData.requestedBy// "Ashish",
        //  Subcategory: "Access",
        },
        subscriptionId: props.subscriptionId,
      };
      const response = await fetch(`${props.Apilink}`, {
        method: "POST",
        headers: {
          "Content-Type": "application/json",
          "Ocp-Apim-Subscription-Key": `${props.OcpApimKey}`,
          Email:formData.requestedFor,// "pmishra@mcit.gov.qa",
        },
        body: JSON.stringify(payload),
      });
      console.log("response", response);
      if (!response.ok) {
        const errorText = await response.text();
        throw new Error(`Request failed: ${response.status} - ${errorText}`);
      }
       if (response.ok) {
        const rawResponse = await response.text();
        const jsonStart = rawResponse.indexOf("{");
        if (jsonStart === -1) {
          throw new Error("JSON not found in response");
        }

        // Step 2: Extract only the JSON string
        const jsonString = rawResponse.slice(jsonStart);

        // Step 3: Parse the JSON
        let parsedData;
        try {
          parsedData = JSON.parse(jsonString);
        } catch (e) {
          throw new Error("Failed to parse JSON: " + e.message);
        }

        const requestRecId = parsedData?.ServiceRequests?.[0]?.strRequestRecId;
        const strRequestNum = parsedData?.ServiceRequests?.[0]?.strRequestNum;

        console.log("requestRecId submitted Hardware Request:", requestRecId);
        console.log("strRequestNum submitted Hardware Request:", strRequestNum);
        let flag = true;
        if (formData.files.length > 0) {
          flag = false;
          await saveRequestAttachment(
            requestRecId,
            strRequestNum,
            formData.files
          );
        }

        if (flag) {
          setModalHeading("Success");
          setModalMessage("Your Request has been submitted successfully.");
          setAlertsection("Accepted");
          setIconLoad("SkypeCircleCheck");
          handleShowModal();

          if (props.isredirect) {
            setTimeout(() => {
              window.location.reload();
            }, 2000);
          }
        }
      }
      
    } catch (error: any) {
      console.error("Error submitting Request:", error);
      setModalHeading("Error");
      setModalMessage(error.message);
      setAlertsection("rejected");
      setIconLoad("ErrorBadge");
      handleShowModal();
    }
  };
  const saveRequestAttachment = async (
    recid: string,
    requestnum: string,
    formData: any
  ) => {
    try {
      console.log("Attachment function is called");
      const ApiformData = new FormData();
      ApiformData.append("ObjectID", recid);
      ApiformData.append("ObjectType", "ServiceReq#");
      ApiformData.append("File", formData[0].content);
      const response = await fetch(props.attachmentApilink, {
        method: "POST",
        headers: {
          "Ocp-Apim-Subscription-Key": props.OcpApimKey, // "ba47658772b3473cbd9eb045e856e9fc",
        },
        body: ApiformData,
      });
      if (response.ok) {
        setModalHeading("Success");
        setModalMessage("Your Request has been submitted successfully.");
        setAlertsection("Accepted");
        setIconLoad("SkypeCircleCheck");
        handleShowModal();

        if (props.isredirect) {
          setTimeout(() => {
            window.location.reload();
          }, 2000);
        }
      }
      if (!response.ok) {
        const errorText = await response.text();
        throw new Error(`Request failed: ${response.status} - ${errorText}`);
      }
    } catch (error: any) {
      console.error("Error submitting Attachment:", error);
      setModalHeading("Error");
      setModalMessage(error.message);
      setAlertsection("rejected");
      setIconLoad("ErrorBadge");
      handleShowModal();
    }
  };
  if (isLoadingUser) {
    return <div>Loading user information...</div>;
  }
  return (
    <>
      <ServiceUIForm
        OcpApimKey={props.OcpApimKey}
        UserRecIdApilink={props.UserRecIdApilink}
        context={props.context}
        userprofileAD={userProfileAD}
        EmpId={userProfileAD?.employeeId || ""}
        onErrorRequiredFields={() => showErrorModal()}
        onSave={async (formData) => {
          await saveRequest(formData);
        }}
      />

      <AlertModal
        showModal={showModal}
        handleShowModal={handleShowModal}
        handleCloseModal={handleCloseModal}
        heading={modalHeading}
        message={modalMessage}
        style={""}
        section={alertsection}
        icon={iconLoad}
      />
    </>
  );
};

export default ServiceRequest;
