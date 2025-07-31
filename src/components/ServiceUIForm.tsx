import * as React from "react";
import { useState, useRef } from "react";
import {
  TextField,
  PrimaryButton,
  DefaultButton,
  Dropdown,
  DatePicker,
  DayOfWeek,
  mergeStyles,
  defaultDatePickerStrings,
  mergeStyleSets,
} from "@fluentui/react";
import { WebPartContext } from "@microsoft/sp-webpart-base";
import "../css/style.css";
import {
  IUserProfile,
  IServiceRequestFormData,
  IServiceProps,
} from "../webparts/service/components/IServiceProps";
import {
  IPeoplePickerContext,
  PeoplePicker,
  PrincipalType,
} from "@pnp/spfx-controls-react/lib/PeoplePicker";
import { Col } from "react-bootstrap";

export interface IRequestUIFormProps {
  context: WebPartContext;
  userprofileAD: IUserProfile;
  EmpId: string;
  onErrorRequiredFields: () => void;
  onSave: (formData: IServiceRequestFormData) => Promise<void>;
}
const isAr =
  window.location.pathname.includes("/ar/") ||
  window.location.search.includes("lang=ar");

const ServiceUIForm: React.FC<IRequestUIFormProps> = (props) => {
  const [formData, setFormData] = useState<IServiceRequestFormData>({
    requestedBy: props.userprofileAD?.displayName,
    requestedFor: "",
    requestedFor_Title: "",
    serviceName: "",
    serviceName_key: "",
    officeLocation: null,
    PhoneNumber: "",
    AccessthroughVPN: "",
    files: [],
    description: "",
    AccessDurationFrom: null,
    AccessDurationTo: null,
    requestedFor_key: "",
    AccessDurationFroms: "",
    AccessDurationTos: "",
  });
  const [uploadedFiles, setUploadedFiles] = useState<Array<{ name: string }>>(
    []
  );
  const [showErrorUpload, setShowErrorUpload] = useState("");
  const [errors, setErrors] = useState<{ [field: string]: string }>({});
  const inputRef = useRef<HTMLInputElement>(null);
  const [firstDayOfWeek, setFirstDayOfWeek] = React.useState(DayOfWeek.Sunday);
  const [selectedPeoplePickerProfiles, setselectedPeoplePickerProfiles] =
    useState<IUserProfile[]>([]);
  const [, setForceUpdater] = useState(0);

  let fileInfo: HTMLInputElement;
  function handleInputChange(field: string, value: any) {
    if (field == "AccessDurationFrom") {
      const dateStr = value.toString(); //
      const parts = dateStr.split(" "); // Split by space
      const monthMap = {
        Jan: "01",
        Feb: "02",
        Mar: "03",
        Apr: "04",
        May: "05",
        Jun: "06",
        Jul: "07",
        Aug: "08",
        Sep: "09",
        Oct: "10",
        Nov: "11",
        Dec: "12",
      };
      const month = monthMap[parts[1]];
      const day = parts[2];
      const year = parts[3];
      const formattedDate = `${month}/${day}/${year}`;
      console.log(formattedDate); // Output: "07/24/2025"
      let col_vals = formattedDate;
      setFormData((prev) => ({ ...prev, AccessDurationFroms: col_vals }));
    }
    if (field == "AccessDurationTo") {
      const dateStr = value.toString(); //
      const parts = dateStr.split(" "); // Split by space
      const monthMap = {
        Jan: "01",
        Feb: "02",
        Mar: "03",
        Apr: "04",
        May: "05",
        Jun: "06",
        Jul: "07",
        Aug: "08",
        Sep: "09",
        Oct: "10",
        Nov: "11",
        Dec: "12",
      };
      const month = monthMap[parts[1]];
      const day = parts[2];
      const year = parts[3];
      const formattedDate = `${month}/${day}/${year}`;
      console.log(formattedDate); // Output: "07/24/2025"
      let col_vals = formattedDate;
      setFormData((prev) => ({ ...prev, AccessDurationTos: col_vals }));
    }
    setFormData((prev) => ({ ...prev, [field]: value }));
  }

  function validateForm() {
    const newErrors: { [field: string]: string } = {};
    //if (!formData.requestedBy.trim()) newErrors.requestedBy = isAr ? "مطلوب بواسطة" : "requestedBy is required";
    if (!formData.requestedFor.trim())
      newErrors.requestedFor = isAr
        ? "مطلوب مطلوب"
        : "RequestedFor is required";
    if (!formData.serviceName.trim())
      newErrors.serviceName = isAr
        ? "اسم الخدمة مطلوب"
        : "Service Name is required";
    if (!formData.officeLocation)
      newErrors.officeLocation = isAr ? "الموقع مطلوب" : "Location is required";
    if (!formData.PhoneNumber)
      newErrors.PhoneNumber = isAr
        ? "رقم الهاتف مطلوب"
        : "Phone Number is required";
    if (formData.serviceName == "VPN Access Request") {
      if (!formData.AccessDurationFrom)
        newErrors.AccessDurationFrom = isAr
          ? "مطلوب مدة الوصول من"
          : "Access Duration From is required";
      if (!formData.AccessDurationTo)
        newErrors.AccessDurationTo = isAr
          ? "مدة الوصول إلى مطلوبة"
          : "Access Duration To is required";
      if (!formData.AccessthroughVPN)
        newErrors.AccessthroughVPN = isAr
          ? "يجب تقديم طلب الوصول إلى الخدمات عبر VPN"
          : "Services Request to Access through VPN is required";
    }

    if (!formData.description)
      newErrors.description = isAr ? "الوصف مطلوب" : "Description is required";
    setErrors(newErrors);

    if (Object.keys(newErrors).length > 0) {
      props.onErrorRequiredFields();
      return false;
    }
    return true;
  }

  async function handleSubmit() {
    //   e.preventDefault();
    setErrors({});
    if (!validateForm()) return;
    await props.onSave(formData);
  }
  const displayName = props.userprofileAD?.displayName;

  let initials = "";
  if (displayName && displayName.trim()) {
    const parts = displayName.split(" ");
    initials = parts[0][0] + parts[parts.length - 1][0];
  } else {
    initials = "";
  }

  const peoplePickerContext: IPeoplePickerContext = {
    absoluteUrl: props.context.pageContext.web.absoluteUrl,
    msGraphClientFactory: props.context.msGraphClientFactory as any,
    spHttpClient: props.context.spHttpClient as any,
  };
  const _getPeoplePickerItems = async (
    selectedUserProfiles: any[],
    internalName: string,
    internalName_text: string
  ) => {
    if (selectedUserProfiles.length > 0) {
      const emails = selectedUserProfiles[0].id.split("|")[2];
      const title = selectedUserProfiles[0].text;
      handleInputChange(internalName, emails);
      handleInputChange(internalName_text, title);
      console.log("Selected userids:", emails);
      console.log("Selected Items:", selectedUserProfiles);
    } else {
      handleInputChange(internalName, "");
      handleInputChange(internalName_text, "");
    }
  };

  let requesterFileList: FileList | null = null;
  const removeAttachment = (fileName: string) => {
    // Filter out the file to remove
    const updatedFile = uploadedFiles.filter((file) => file.name !== fileName);

    // Update the state with the new list of files
    setUploadedFiles(updatedFile);
    handleInputChange("files", updatedFile);
    // Update the formData to reflect the removal
  };

  const readFile = (e: React.ChangeEvent<HTMLInputElement>, field) => {
    requesterFileList = e.target.files;
    if (requesterFileList) {
      console.log("file details", fileInfo.files[0]);
      const fileExtension = fileInfo.files[0].name.substring(
        fileInfo.files[0].name.lastIndexOf(".") + 1,
        fileInfo.files[0].name.length
      );
      const fileName =
        fileInfo.files[0].name
          .substring(0, fileInfo.files[0].name.lastIndexOf(".") + 1)
          .replace(/[&\/\\#~%":*. [\]!¤+`´^?<>|{}]/g, "") +
        "." +
        fileExtension;

      const newFile = {
        name: fileName,
        content: fileInfo.files[0],
      };

      // Add the new file to the existing state of uploaded files
      setUploadedFiles((prevFiles) => {
        const updatedFiles = [...prevFiles, newFile];
        console.log("uploadedFiles file details", updatedFiles);

        // Update formData using the latest updatedFiles
        setFormData((prev) => ({ ...prev, [field]: updatedFiles }));

        return updatedFiles;
      });
      // Update progress for the newly added file
      let currentProgress = 0;
      const interval = setInterval(() => {
        if (currentProgress >= 100) {
          clearInterval(interval);
        } else {
          currentProgress += 10;
          setUploadedFiles((prevFiles) =>
            prevFiles.map((file) =>
              file.name === newFile.name
                ? { ...file, progress: currentProgress }
                : file
            )
          );
        }
      }, 300);
    }
  };

  const updateFormData = (
    event: React.FormEvent<HTMLInputElement | HTMLTextAreaElement>,
    newValue: string | undefined,
    column: any
  ) => {
    // newValue is the updated text from the Fluent UI TextField
    const value = newValue ?? "";

    // Update formData
    setFormData((prev) => ({
      ...prev,
      [column]: value,
    }));

    // Remove the field's error if the user typed something valid
    setErrors((prevErrors) => {
      const newErrors = { ...prevErrors };
      if (newErrors[column] && value.trim() !== "") {
        delete newErrors[column];
      }
      return newErrors;
    });

    forceUpdate();
  };
  const updateFormDropData = (option: any, column: any, columnKey: any) => {
    setFormData((prev) => ({ ...prev, [columnKey]: option?.key as string }));
    setFormData((prev) => ({ ...prev, [column]: option?.text as string }));

    setErrors((prevErrors) => {
      const newErrors = { ...prevErrors };
      if (newErrors[column] && option.key) {
        delete newErrors[column];
      }
      return newErrors;
    });

    forceUpdate();
  };
  const forceUpdate = () => setForceUpdater((prev) => prev + 1);
  const _getPeoplePickerMemberItems = async (
    selectedUserProfiles: any[],
    Member: string
  ) => {
    if (selectedUserProfiles.length > 0) {
      const emails = selectedUserProfiles[0].id.split("|")[2];
      handleInputChange(Member, emails);
      console.log("Selected userids:", emails);
    } else {
      handleInputChange(Member, "");
    }
  };
  return (
    <div>
      <div className="maincontainer">
        <div className="header-top">
          <div className="person-image">{initials}</div>
          <div>
            <div className="person-name">
              {props.userprofileAD?.displayName}
            </div>
            <div className="person-description">
              {props.userprofileAD?.jobTitle} | ID:{" "}
              {props.EmpId ? props.EmpId : "N/A"}
            </div>
          </div>
        </div>
        <div className="textContainer">
          <h2 className="form-heading">
            {isAr ? "يرجى ملء النموذج أدناه" : "Please fill up the form below"}
          </h2>

          <div className="fieldContainer">
            {/* Requested By */}
            <TextField
              type="text"
              label={isAr ? "تم الطلب بواسطة" : "Requested By"}
              className="form-text"
              readOnly
              value={props.userprofileAD?.displayName}
            />
            <div
              className={`people-picker-wrapper ${
                errors.requestedFor ? "error-border" : ""
              }`}
            >
              <PeoplePicker
                context={peoplePickerContext}
                titleText={isAr ? "مطلوب ل *" : "Requested for *"}
                personSelectionLimit={1}
                groupName={""}
                defaultSelectedUsers={[formData.requestedFor]}
                showtooltip={true}
                disabled={false}
                searchTextLimit={3}
                onChange={(e) => {
                  _getPeoplePickerItems(e, "requestedFor","requestedFor_Title");
                }}
                principalTypes={[PrincipalType.User]}
                resolveDelay={1000}
              />
            </div>

            <TextField
              label={isAr ? "موقع *" : "Location *"}
              value={formData.officeLocation}
              onChange={(ev, newValue) =>
                updateFormData(ev, newValue, "officeLocation")
              }
              className={`form-text  ${
                errors.officeLocation ? "error-field" : ""
              }`}
            />

            <TextField
              label={isAr ? "رقم التليفون *" : "Phone Number *"}
              value={formData.PhoneNumber}
              className={`form-text  ${
                errors.PhoneNumber ? "error-field" : ""
              }`}
              onChange={(ev, newValue) => {
                // Allow empty string or digits only
                if (newValue === "" || /^\d+$/.test(newValue)) {
                  updateFormData(ev, newValue, "PhoneNumber");
                }
              }}
              inputMode="numeric"
            />

            <Dropdown
              label={isAr ? "اسم الخدمة *" : "Service Name *"}
              selectedKey={formData.serviceName_key}
              className={`dropdownfield ${
                !formData.serviceName ? "placeholder-gray" : ""
              } ${errors.serviceName ? "error-field" : ""}`}
              styles={{
                dropdown: {
                  borderColor: errors.serviceName ? "red" : undefined,
                },
              }}
              onChange={(_, option) => {
                updateFormDropData(option, "serviceName", "serviceName_key");
              }}
              options={[
                {
                  key: "",
                  text: isAr ? "حدد اسم الخدمة..." : "Select Service Name...",
                  disabled: true,
                },
                {
                  key: "89E528D6F35E4515B9076B9F0BB4508D",
                  text: isAr ? "طلب الوصول إلى VPN" : "VPN Access Request",
                },
                {
                  key: "76C6CBEA950A46E5AAB7CE8E93EA39F4",
                  text: isAr ? "إعادة تعيين VPN MFA/OTP" : "Reset VPN MFA/OTP",
                },
                {
                  key: "C3A24044A1CE44088F3425504692B631",
                  text: isAr ? "أخرى" : "Others",
                },
              ]}
            />
            {formData.serviceName == "VPN Access Request" && (
              <>
                <DatePicker
                  label={isAr ? "مدة الوصول من *" : "Access Duration From *"}
                  firstDayOfWeek={firstDayOfWeek}
                  placeholder="Select a date..."
                  ariaLabel="Select a date"
                  onSelectDate={(e): void => {
                    handleInputChange("AccessDurationFrom", e);
                  }}
                  className={`form-text ${
                    errors.AccessDurationFrom ? "has-error" : "no-error"
                  }`}
                  // DatePicker uses English strings by default. For localized apps, you must override this prop.
                  strings={defaultDatePickerStrings}
                />
                <DatePicker
                  label={isAr ? "مدة الوصول إلى *" : "Access Duration To *"}
                  firstDayOfWeek={firstDayOfWeek}
                  placeholder="Select a date..."
                  ariaLabel="Select a date"
                  onSelectDate={(e): void => {
                    handleInputChange("AccessDurationTo", e);
                  }}
                  className={`form-text ${
                    errors.AccessDurationTo ? "has-error" : "no-error"
                  }`}
                  // DatePicker uses English strings by default. For localized apps, you must override this prop.
                  strings={defaultDatePickerStrings}
                />
                <TextField
                  label={
                    isAr
                      ? "طلب الوصول إلى الخدمات عبر VPN *"
                      : "Services Request to Access through VPN *"
                  }
                  value={formData.AccessthroughVPN}
                  className={`form-text  ${
                    errors.AccessthroughVPN ? "error-field" : ""
                  }`}
                  onChange={(ev, newValue) => {
                    updateFormData(ev, newValue, "AccessthroughVPN");
                  }}
                />
              </>
            )}
          </div>
          <div
            className="description_div"
            style={{
              marginTop:
                formData.serviceName == "" ||
                formData.serviceName == "Reset VPN MFA/OTP" ||
                formData.serviceName == "Others"
                  ? "80px"
                  : "",
            }}
          >
            <TextField
              label={isAr ? "وصف *" : "Description *"}
              value={formData.description}
              multiline
              rows={4}
              type="text-area"
              className={`text-area ${errors.description ? "error-field" : ""}`}
              onChange={(ev, newValue) =>
                updateFormData(ev, newValue, "description")
              }
              styles={{
                root: { color: "#555" },
                fieldGroup: { border: "1px solid #ccc" },
                field: { color: "#555" },
              }}
            />
          </div>
          <Col className="mt-4">
            <div style={{ display: "flex", alignItems: "end" }}>
              <label
                style={{
                  marginRight: "4px",
                  marginTop: "24px",
                  fontSize: "12px",
                  fontFamily: "Segoe UI",
                  color: "#555555",
                  fontWeight: "500",
                  marginBottom: "11.5px",
                }}
              >
                {isAr
                  ? "أي مستندات أو صور تساعد في إثبات القضية:"
                  : "Any documents or pictures (optional):"}
              </label>
            </div>

            <div className="attachment-container">
              <div className="attachment-placeholder">
                <img
                  className="attachment-icon"
                  src="data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAACgAAAAoCAYAAACM/rhtAAAACXBIWXMAAAsTAAALEwEAmpwYAAAAAXNSR0IArs4c6QAAAARnQU1BAACxjwv8YQUAAAKoSURBVHgB7ZnbjdpAFIaPDYLX7WDdQUBcJJ4CFWRTQUgHbAVABbAVQCpIOljyhIS4leAO4rwB4pL/sMcbxxdij8dWFO0vjcY7c8Z8mjOeOXPWoAgtFotKsVgc4bGCckfpZO/3+06r1bIpoYywxuVy+WCa5kTAHJSfpK57qZUgA4Cbzca6XC7PeOT6CaCDarXqkKLW67WdBtL0N5zP567ATWu1Wi8NnE9bfm+5XH7GJMReMgFAwzDeXztM8wtpFN7bIYFkD8WFNKM6drudTRrFnhBIG6USF9KkHBUC+fVvY3IFZAHS9kC2V6vV5JZ9boDz+dxyn72QqLu3IDMHhBu/cV0qlfre9hDIUdj4ImUs7AZjQH5gCOyJbXpx7VVof7VDfw+QDra24R/jKWN5Zuo7ioXS9hWLfkN2/eMzn0EWQzKMbCuVMBs5vQLSCggX9VDdw02PYf1yKs3C+uD+sGZ9LgZcHy4ayVqakCZpARS4gfu3fJV90qDUgD44N7DgE2OgAzIVIOLGTy4cFvlnkrjxeDx2dEEqA0pQOxW4R3wYU7ev2WxuT6fTR34WyB4pSgmQrwMScTPcEHBjv02j0ZjJrDLkiGebFKQEWCgUHlDdCdwgyo5nlWeXXiCrpCAlQIbiH7wF57Ed80nCVwdSkPJGjU13m8B2RorKPR5MqjfAtPrnAXWHWz9QLqRRWgElMNUqrYAasxCv+r/WoBz67yiFcPRt6/X6U1z7RIBYY3zgVyiF8A4bVWaAHEJZlE6xj0hWIkC5ndmUo8I+kmtUjPSvRTkJ11F32QR2gQAgFvGMa8R8Wi49cYTf7EkdcH8AUMJ4myTzxNEzZSROKElOhj8+53A4DP02oUl0b56a8pHDFy2+y/g7jKgRAjkgPf+GiBRnvzBz46jE+i8JiDR7F2tlUAAAAABJRU5ErkJggg=="
                  alt="Attachment Icon"
                />
                {isAr
                  ? "إرفاق الملف بتنسيق PNG أو JPG أو PDF (اختياري)"
                  : "Attach file in PNG, JPG, or PDF format (optional)"}
                <input
                  type="file"
                  // ref={inputRef}
                  multiple={true}
                  ref={(element) => {
                    fileInfo = element;
                  }}
                  onChange={(e) => {
                    readFile(e, "files");
                  }}
                />
              </div>
              <span style={{ color: "red" }}>
                {errors.files || showErrorUpload}
              </span>
            </div>

            {uploadedFiles.map((file, index) => (
              <div key={index}>
                <div className="uploadeditems">
                  <strong>{file.name}</strong>
                  <div className="progresscontainer">
                    <div
                      className="progressbar"
                      id="progressbar"
                      style={{ width: `${file["progress"]}%` }} // Each file has its own progress
                    ></div>
                  </div>
                  <div
                    className="cancelbtn"
                    onClick={() => {
                      removeAttachment(file.name); // Pass the file name to remove it
                    }}
                  >
                    X
                  </div>
                </div>
              </div>
            ))}
            {/* <p style={{ color: "gray" }}>
                  {!isAr
                    ? "# You can upload up to 10 documents or images."
                    : "يمكنك تحميل ما يصل إلى 10 مستندات أو صور."}
                </p> */}
          </Col>
          <div className="buttonContainer">
            <PrimaryButton
              onClick={() => {
                handleSubmit();
              }}
              styles={{ root: { fontSize: "20px" } }}
              text={!isAr ? "Submit" : "يُقدِّم"}
              className="submit-formbtn"
            />
            <DefaultButton
              text={!isAr ? "Cancel" : "يلغي"}
              className="cancel-formbtn"
              onClick={() => {
                setFormData({
                  requestedFor: "",
                  requestedFor_Title: "",
                  serviceName: "",
                  serviceName_key: "",
                  officeLocation: "",
                  PhoneNumber: "",
                  AccessthroughVPN: "",
                  description: "",
                  files: [],
                  AccessDurationFrom: null,
                  AccessDurationTo: null,
                  requestedFor_key: "",
                  AccessDurationFroms: "",
                  AccessDurationTos: "",
                });
                setUploadedFiles([]);
                setShowErrorUpload("");
                setErrors({});
                if (inputRef.current) inputRef.current.value = "";
                fileInfo = null;
              }}
            />
          </div>
        </div>
      </div>
      <div className="testelement"></div>
    </div>
  );
};

export default ServiceUIForm;
