import React, { Component } from "react";
import Dropzone from "react-dropzone";
import { IInputs } from "../generated/ManifestTypes";
import "../css/Dropzone.css";
import {
  createRelatedNote,
  getRelatedNotes,
  updateRelatedNote,
  deleteRelatedNote,
  getSharePointLocations,
  getSharePointFolderData,
  createSharePointDocument,
  deleteSharePointDocument,
  createSharePointFolder,
  createSharePointLocation,
  createAtivityDocument,
  getNoteViews
} from "../DataverseActions";
import {
  FileData,
  SharePointDocument,
  PreviewFile,
  GenericActionResponse,
  NoteView,
} from "../Interfaces";
import { Img } from "react-image";
import {
  isPDF,
  isImage,
  createDataUri,
  isActivityType,
  focusSPDocumentsAndRestore,
  getLocalString,
  b64toBlob,
} from "../utils";
import {
  IContextualMenuItem,
  DefaultButton,
  PrimaryButton,
  TextField,
  Dialog,
  DialogType,
  DialogFooter,
  Stack,
  IStackStyles,
  IStackTokens,
  CommandButton,
  SearchBox,
  Spinner,
  SpinnerSize,
  Toggle,
  TooltipHost,
  IButtonStyles,
  Dropdown,
  DropdownMenuItemType,
  IDropdownStyles,
  IDropdownOption,
  IconButton,
  Icon,
  ITooltipHostStyles,
  ContextualMenu,
  IContextualMenuProps,
  ContextualMenuItemType,
  Modal,
  IFocusTrapZoneProps,
  IIconProps,
  Callout,
  Label,
  HighContrastSelector,
  DirectionalHint,
  ComboBox,
  IComboBox,
  IComboBoxOption,
  IButtonProps,
} from "@fluentui/react";
import {
  getFileTypeIconProps,
  initializeFileTypeIcons,
} from "@uifabric/file-type-icons";
import { Tooltip } from "react-tippy";
import "react-tippy/dist/tippy.css";
import toast, { Toaster } from "react-hot-toast";
import FolderIcon from "./FolderIcon";
import { getEntityMetadata, getControlValue } from "../utils";
import ReactViewer from 'react-viewer';
import { LocalStrings } from "../consts/LocalStrings";

export interface LandingProps {
  context?: ComponentFramework.Context<IInputs>;
  isDisabled: boolean;
}

interface LandingState {
  noteViews: NoteView[];
  selectedViewId?: string;
  isViewsLoading: boolean;
  files: FileData[];
  editingFileId?: string;
  selectedFiles: string[];
  spSearchText: string;
  notesSearchText: string;
  sortAsc: boolean;
  isLoading: boolean;
  sharePointDocLoc: boolean;
  showTooltip: boolean;
  documentLocations: {
    name: string;
    sharepointdocumentlocationid: string;
    parentsiteid?: string;
  }[];
  selectedDocumentLocation: string | null;
  sharePointData: SharePointDocument[];
  currentFolderPath: string;
  folderStack: string[];
  showModal: boolean;
  newFolderName: string;
  selectedDocumentLocationName: string;
  isCollapsed: boolean;
  menuVisible: boolean;
  target: HTMLElement | null;
  previewFile: PreviewFile | null;
  isDialogOpen: boolean;
  xlsxContent: string;
  xlsxData: SheetData | null;
  sharePointEnabled: boolean;
  userPreference: boolean | undefined;
  isCalloutVisible: boolean;
  isFolderDeletionDialogVisible: boolean;
  isCreateLocationDialogVisible: boolean;
  createLocationDisplayName: string;
  selectedCreateLocation: string;
  createLocationFolderName: string;
  isSaveButtonEnabled: boolean;
  isActivity: boolean;
  sharePointEnabledParameter: boolean;
  selectedFolderForDelete: SharePointDocument | null;
  formType: number;
}

type FileAction =
  | "edit"
  | "download"
  | "delete"
  | "preview"
  | "addToActivityAttachment";

const SharePointDocLocTooltip = (
  <span>
    Upload files directly to a SharePoint Document Location. To learn more go to{" "}
    <a
      href="https://learn.microsoft.com/en-us/power-platform/admin/create-edit-document-location-records"
      target="_blank"
      rel="noopener noreferrer"
    >
      MS Docs
    </a>
    .
  </span>
);
const tooltipId = "sharePointDocLocTooltip";
const buttonStyles: Partial<ITooltipHostStyles> = {
  root: {
    border: "none",
    minWidth: "auto",
    margin: 0,
    padding: 0,
    display: "inline-block",
  },
};
const dropdownStyles: Partial<IDropdownStyles> = {
  dropdown: { width: 200 },
};
const ribbonStyles: IStackStyles = {
  root: {
    alignItems: "center",
    display: "flex",
    width: "100%",
    justifyContent: "space-between",
  },
};

const searchRibbonStyles: IStackStyles = {
  root: {
    marginBottom: "0px",
  },
};

type SheetData = {
  [sheetName: string]: any[];
};

const isValidBase64 = (str: string) => {
  try {
    return btoa(atob(str)) === str;
  } catch (err) {
    return false;
  }
};

const ribbonStackTokens: IStackTokens = { childrenGap: 10 };

const EXCLUDED_NOTE_VIEWS = new Set<string>([
  "Notes Associated View",
  "Notes",
  "Notes Advanced Find View",
  "Notes Lookup View",
  "Quick Find Annotations",
]);

export class Landing extends Component<LandingProps, LandingState> {
  private targetRef: React.RefObject<HTMLDivElement>;
  private viewerHost = React.createRef<HTMLDivElement>();
  constructor(props: LandingProps) {
    super(props);
    initializeFileTypeIcons();
    this.state = {
      files: [],
      noteViews: [],
      selectedViewId: undefined,
      isViewsLoading: false,
      editingFileId: undefined,
      selectedFiles: [],
      spSearchText: "",
      notesSearchText: "",
      sortAsc: true,
      isLoading: true,
      // CUSTOM 2026-04-08: sharePointDocLoc starts as false; componentDidMount forces
      // it to true when sharePointOnlyMode is enabled (see componentDidMount).
      sharePointDocLoc: false,
      showTooltip: false,
      documentLocations: [],
      selectedDocumentLocation: null,
      sharePointData: [],
      currentFolderPath: "",
      folderStack: [],
      showModal: false,
      newFolderName: "",
      selectedDocumentLocationName: "",
      isCollapsed: window.innerWidth < 767,
      menuVisible: false,
      target: null,
      previewFile: null,
      isDialogOpen: false,
      xlsxContent: "",
      xlsxData: null,
      sharePointEnabled: false,
      userPreference: false,
      isCalloutVisible: false,
      isFolderDeletionDialogVisible: false,
      isCreateLocationDialogVisible: false,
      createLocationDisplayName: "",
      selectedCreateLocation: "",
      createLocationFolderName: "",
      isSaveButtonEnabled: false,
      isActivity: false,
      sharePointEnabledParameter: false,
      selectedFolderForDelete: null,
      formType: 0,
    };
    this.removeFile = this.removeFile.bind(this);
    this.downloadFile = this.downloadFile.bind(this);
    this.toggleSharePointDocLoc = this.toggleSharePointDocLoc.bind(this);
    this.toggleTooltip = this.toggleTooltip.bind(this);
    this.getSharePointLocations = this.getSharePointLocations.bind(this);
    this.getSharePointData = this.getSharePointData.bind(this);
    this.handleFolderClick = this.handleFolderClick.bind(this);
    this.handleFolderClick = this.handleFolderClick.bind(this);
    this.handleBackClick = this.handleBackClick.bind(this);
    this.handleDropdownChange = this.handleDropdownChange.bind(this);
    this.handleResize = this.handleResize.bind(this);
    this.toggleMenu = this.toggleMenu.bind(this);
    this.closeMenu = this.closeMenu.bind(this);
    this.openDialog = this.openDialog.bind(this);
    this.closeDialog = this.closeDialog.bind(this);
    this.onGearIconClick = this.onGearIconClick.bind(this);
    this.onCalloutDismiss = this.onCalloutDismiss.bind(this);
    this.toggleRememberLocation = this.toggleRememberLocation.bind(this);
    this.handleRemoveFolderClick = this.handleRemoveFolderClick.bind(this);
    this.isActivityType = this.isActivityType.bind(this);
    this.addFileAttachmentToActivity =
      this.addFileAttachmentToActivity.bind(this);
    this.targetRef = React.createRef<HTMLDivElement>();
  }

  delay(ms: number) {
    return new Promise((resolve) => setTimeout(resolve, ms));
  }

  async addFileAttachmentToActivity(file: SharePointDocument): Promise<void> {
    const { context } = this.props;
    const entityId = (context as any).page.entityId;
    const entityTypeName = (context as any).page.entityTypeName;
    if (!entityId || !entityTypeName) {
      toast.error("Missing entity info for activity attachment.");
      return;
    }
    if (!file.absoluteurl) {
      toast.error("Missing URL for SharePoint file.");
      return;
    }
    try {
      const res = await createAtivityDocument(context!, file.absoluteurl, entityId, entityTypeName);
      if (res.success) {
        toast.success(`${file.fullname} attached to activity.`);
      } else {
        toast.error(`Failed to attach: ${res.message}`);
      }
    } catch (err) {
      toast.error("Unexpected error attaching file.");
    }
  }

  onGearIconClick = () => {
    this.setState((prev) => ({ isCalloutVisible: !prev.isCalloutVisible }));
  };

  onCalloutDismiss = () => {
    this.setState({ isCalloutVisible: false });
  };

  async getSharePointLocations(): Promise<void> {
    const response = await getSharePointLocations(this.props.context!);

    if (response.length === 0) {
      this.setState({
        documentLocations: [],
        selectedDocumentLocation: null,
        selectedDocumentLocationName: "",
        sharePointData: [],
        currentFolderPath: "",
        folderStack: [],
      });
      return;
    }

    this.setState({
      documentLocations: response,
      selectedDocumentLocation: response[0].sharepointdocumentlocationid,
      selectedDocumentLocationName: response[0].name,
    });
  }

  private isDefaultLocation() {
    if (
      this.state.documentLocations.length === 1 &&
      this.state.selectedDocumentLocationName === "Documents on Default Site 1"
    ) {
      return false;
    } else {
      return true;
    }
  }
  private async getSharePointData(popFolderStack = false): Promise<void> {
    if (!this.state.selectedDocumentLocation) {
      this.setState({ sharePointData: [], isLoading: false });
      return;
    }
    this.setState(
      prev => {
        const folderStack = [...prev.folderStack];
        const currentFolderPath = popFolderStack
          ? folderStack.pop() || ""
          : prev.currentFolderPath || "";
        return { folderStack, currentFolderPath };
      },
      async () => {
        const { currentFolderPath } = this.state;
        const defaultLocation = this.isDefaultLocation();
        const data = await getSharePointFolderData(
          this.props.context!,
          currentFolderPath,
          this.state.selectedDocumentLocation,
          this.state.selectedDocumentLocationName,
          defaultLocation
        );
        this.setState({ sharePointData: data });
      }
    );
  }

  async isActivityType(): Promise<boolean> {
    const metadata = await getEntityMetadata(this.props.context!);
    if (!metadata) {
      return false;
    }
    const isActivity = isActivityType(metadata.schemaName);
    if (isActivity) {
      return true;
    } else {
      return false;
    }
  }

  async getUserSettings(): Promise<any> {
    const settings = localStorage.getItem("userSettings");
    return settings ? JSON.parse(settings) : null;
  }

  async checkUserSettings(): Promise<boolean | undefined> {
    const settings = localStorage.getItem("userSettings");

    if (settings) {
      const parsedSettings = JSON.parse(settings);
      const metadata = await getEntityMetadata(this.props.context!);
      if (
        parsedSettings.tableName !== metadata?.schemaName ||
        parsedSettings.entityId !== metadata?.entityId
      ) {
        return undefined;
      }
      if (parsedSettings.tableName === metadata?.schemaName) {
        if (typeof parsedSettings.NotesOrSharePoint === "undefined") {
          return parsedSettings.NotesOrSharePoint;
        } else {
          return undefined;
        }
      } else {
        return false;
      }
    } else {
      return false;
    }
  }

  async checkSharePointIntegration(): Promise<boolean> {
    const response = await getSharePointLocations(this.props.context!);
    return response.length === 0 ? true : false;
  }

  async componentDidMount() {
    this.setState({ isViewsLoading: true });

    const defaultNotesViewName =
      (getControlValue(this.props.context!, "defaultNotesView") as string) ?? ""
        .trim();

    try {
      const res = await getNoteViews(this.props.context!);

      if (res.success) {
        let availableViews = res.data.filter(
          (v) => !EXCLUDED_NOTE_VIEWS.has(v.name)
        );

        const allNotes = availableViews.find(
          (v) => v.name.toLowerCase() === "all notes"
        );
        const otherViews = availableViews.filter(
          (v) => v.name.toLowerCase() !== "all notes"
        );
        availableViews = allNotes ? [allNotes, ...otherViews] : availableViews;

        const findByName = (name: string) =>
          availableViews.find((v) => v.name.toLowerCase() === name.toLowerCase());

        const defaultView =
          (defaultNotesViewName && findByName(defaultNotesViewName)) ||
          findByName("All Notes") ||
          availableViews[0];

        this.setState({
          noteViews: availableViews,
          selectedViewId: defaultView?.savedqueryid,
          isViewsLoading: false,
        });
      } else {
        console.warn(res.message);
        this.setState({ isViewsLoading: false });
      }
    } catch (err) {
      console.error("Failed to load note views:", err);
      this.setState({ isViewsLoading: false });
    }

    var formType = Xrm.Page.ui.getFormType();
    this.setState({ formType });

    // Onsave handler, works well with Active and Deactive ribbon buttons
    Xrm.Page.data.entity.addOnSave(() => {
      const interval = 500;
      const duration = 5000;
      const endTime = Date.now() + duration;

      const intervalId = setInterval(() => {
        const formType = Xrm.Page.ui.getFormType();
        this.setState({ formType });

        if (Date.now() >= endTime) {
          clearInterval(intervalId);
        }
      }, interval);
    });

    window.addEventListener("recordSavedEvent", async (event: any) => {
      const passedEntityId = event.detail.entityId;
      const checkEntityId = async (): Promise<void> => {
        const entityId = (this.props.context as any).page.entityId;
        if (entityId && entityId === passedEntityId) {
          if (this.state.sharePointEnabledParameter == true) {
            focusSPDocumentsAndRestore();
          }
        } else {
          await this.delay(2000);
          return checkEntityId();
        }
      };
      await checkEntityId();
    });

    const sharePointEnabled = await this.checkSharePointIntegration();
    const userPreference = await this.checkUserSettings();
    const settings = await this.getUserSettings();
    const isActivity = await this.isActivityType();

    let sharePointEnabledParameter =
      getControlValue(this.props.context!, "enableSharePointDocuments") === true;
    this.setState({ sharePointEnabledParameter });

    // ============================================================
    // CUSTOM 2026-04-08: sharePointOnlyMode
    // When true, bypass all toggle/preference logic and force
    // SharePoint mode on immediately, skipping Notes entirely.
    // ============================================================
    const sharePointOnlyMode =
      getControlValue(this.props.context!, "sharePointOnlyMode") !== false;
    // NOTE: we default to TRUE if the property is missing/unset (new installs)
    // because default-value="true" in the manifest. We use !== false rather than
    // === true to handle the case where the property hasn't been set yet.

    if (sharePointOnlyMode) {
      // Force SharePoint mode: load locations then switch directly to SP view.
      // We do NOT restore user preferences in this mode.
      await this.getSharePointLocations();
      this.setState(
        {
          sharePointEnabled: sharePointEnabled,
          userPreference: undefined, // not relevant in SP-only mode
          isActivity,
          sharePointDocLoc: true, // always SP mode
        },
        () => {
          this.loadExistingFiles().then(() => {
            this.setState({ isLoading: false });
          });
        }
      );
    } else {
      // ---- Original logic (sharePointOnlyMode = false) ----
      if (
        settings &&
        settings.selectedDocumentLocation &&
        settings.selectedDocumentLocationName &&
        (sharePointEnabledParameter === true || sharePointEnabled === false)
      ) {
        this.setState(
          {
            selectedDocumentLocation: settings.selectedDocumentLocation,
            selectedDocumentLocationName: settings.selectedDocumentLocationName,
          },
          () => {
            this.setState(
              {
                sharePointEnabled: sharePointEnabled,
                userPreference,
                isActivity,
                sharePointDocLoc: true,
              },
              () => {
                this.loadExistingFiles().then(() => {
                  this.setState({ isLoading: false });
                });
              }
            );
          }
        );
      } else {
        this.setState(
          {
            sharePointEnabled: sharePointEnabled,
            userPreference,
            isActivity,
          },
          () => {
            this.loadExistingFiles().then(() => {
              this.setState({ isLoading: false });
            });
          }
        );
      }
      // ---- End original logic ----
    }

    window.addEventListener("resize", this.handleResize);
  }

  componentDidUpdate(prevProps: LandingProps) {
    const prevId = (prevProps.context as any).page.entityId;
    const currId = (this.props.context as any).page.entityId;

    if (prevId !== currId) {
      this.setState({
        files: [],
        sharePointData: [],
        currentFolderPath: "",
        folderStack: [],
        selectedFiles: [],
        isLoading: true,
      }, () => {
        this.loadExistingFiles().then(() => this.setState({ isLoading: false }));
      });
    }
  }

  componentWillUnmount() {
    window.removeEventListener("resize", this.handleResize);
  }

  handleResize() {
    const element = document.querySelector(".ribbon-dropzone-wrapper");
    if (element) {
      const wrapper = element.getBoundingClientRect();
      this.setState({ isCollapsed: wrapper.width < 703 });
    }
  }

  toggleMenu(event: React.MouseEvent<HTMLElement>) {
    this.setState({
      menuVisible: !this.state.menuVisible,
      target: event.currentTarget,
    });
  }

  closeMenu() {
    this.setState({ menuVisible: false });
  }

  getFileExtension(filename?: string): string {
    if (!filename) {
      return "folder";
    }
    const extension = filename.split(".").pop()?.toLowerCase();
    return extension ? extension : "txt";
  }

  handleDrop = async (acceptedFiles: File[]) => {
    const { context, isDisabled } = this.props;
    if (this.state.formType !== 2) {
      return;
    }

    // CUSTOM 2026-04-08: When sharePointOnlyMode is active, override both drop
    // parameters: note drops are always disabled, SP drops are always enabled.
    const sharePointOnlyMode =
      getControlValue(context!, "sharePointOnlyMode") !== false;

    const allowNoteDropsParameter = sharePointOnlyMode
      ? false // Notes drops always disabled in SP-only mode
      : !isDisabled && getControlValue(context!, "allowNoteDrops") === true;

    const allowSharePointDropsParameter = sharePointOnlyMode
      ? !isDisabled // SP drops enabled (subject only to control-disabled flag)
      : !isDisabled && getControlValue(context!, "allowSharePointDrops") === true;

    if (!this.state.sharePointDocLoc && allowNoteDropsParameter) {
      acceptedFiles.forEach((file) => {
        const reader = new FileReader();
        reader.onload = () => {
          const binaryStr = reader.result as string;
          const createNotePromise = createRelatedNote(
            this.props.context!,
            file.name,
            binaryStr,
            file.size,
            file.type
          );

          toast.promise(createNotePromise, {
            loading: getLocalString(context!, LocalStrings.Toast.Message_Uploading_Notes),
            success: (res) => {
              if (res.noteId) {
                this.setState((prevState) => {
                  const newFiles = [
                    ...prevState.files,
                    {
                      filename: file.name,
                      filesize: file.size,
                      documentbody: binaryStr,
                      mimetype: file.type,
                      noteId: res.noteId,
                      createdon: new Date(),
                      subject: "",
                      notetext: "",
                    },
                  ];
                  newFiles.sort(
                    (a, b) => b.createdon.getTime() - a.createdon.getTime()
                  );
                  return { files: newFiles };
                });
                getLocalString(context!, LocalStrings.Button.Label_Preview);
                return `File ${file.name} ${getLocalString(context!, LocalStrings.Toast.Message_Upload_Success_Notes)}`;
              } else {
                throw new Error("Note ID was not returned");
              }
            },
            error: getLocalString(context!, LocalStrings.Toast.Message_Upload_Error_Notes),
          }).then(() => this.loadExistingFiles());
        };

        reader.onerror = () => {
          toast.error(`${getLocalString(context!, LocalStrings.Toast.Message_Read_Error_Notes)} ${file.name}`);
        };

        reader.readAsDataURL(file);
      });
    } else if (allowSharePointDropsParameter) {
      acceptedFiles.forEach((file) => {
        const reader = new FileReader();
        reader.onload = () => {
          const binaryStr = reader.result as string;
          let defaultLocation;
          if (
            this.state.documentLocations.length == 1 &&
            this.state.selectedDocumentLocationName ==
            "Documents on Default Site 1"
          ) {
            defaultLocation = "";
          } else {
            defaultLocation = this.state.selectedDocumentLocation!;
          }

          const uploadPromise = createSharePointDocument(
            this.props.context!,
            file.name,
            binaryStr,
            defaultLocation,
            this.state.currentFolderPath,
            this.state.selectedDocumentLocationName,
            this.isDefaultLocation()
          );

          toast.promise(uploadPromise, {
            loading: getLocalString(context!, LocalStrings.Toast.Message_Uploading_SharePoint),
            success: (res) => {
              if (res.success) {
                this.loadExistingFiles();
                return `File ${file.name} ${getLocalString(context!, LocalStrings.Toast.Message_Upload_Success_SharePoint)}`;
              } else {
                throw new Error(res.message);
              }
            },
            error: getLocalString(context!, LocalStrings.Toast.Message_Upload_Error_SharePoint),
          });
        };

        reader.onerror = () => {
          toast.error(`${getLocalString(context!, LocalStrings.Toast.Message_Read_Error_SharePoint)} ${file.name}`);
        };

        reader.readAsDataURL(file);
      });
    }
  };

  async loadExistingFiles(): Promise<void> {
    const { sharePointDocLoc, selectedDocumentLocation } = this.state;

    if (sharePointDocLoc) {
      await this.getSharePointLocations();
      await this.getSharePointData();
    } else {
      const viewId = this.state.selectedViewId;
      const notes = await getRelatedNotes(this.props.context!, viewId);
      this.setState({ files: notes });
    }
  }

  async removeFile(fileId: string): Promise<void> {
    const { sharePointDocLoc } = this.state;

    if (sharePointDocLoc) {
      const deletePromise = deleteSharePointDocument(
        this.props.context!,
        fileId
      );
      toast.promise(deletePromise, {
        loading: "Deleting...",
        success: () => {
          this.setState((prevState) => ({
            sharePointData: prevState.sharePointData.filter(
              (doc) => doc.sharepointdocumentid !== fileId
            ),
          }));
          return "File deleted.";
        },
        error: "Error deleting file.",
      });
    } else {
      const deletePromise = deleteRelatedNote(this.props.context!, fileId);
      toast.promise(deletePromise, {
        loading: "Deleting...",
        success: () => {
          this.setState((prevState) => ({
            files: prevState.files.filter((file) => file.noteId !== fileId),
          }));
          return "File deleted.";
        },
        error: "Error deleting file.",
      });
    }
  }

  async downloadFile(file: FileData | SharePointDocument): Promise<void> {
    if ("absoluteurl" in file) {
      if (!file.absoluteurl) {
        toast.error(getLocalString(this.props.context!, LocalStrings.Toast.Message_Download_Error_SharePoint));
        return;
      }
      window.open(file.absoluteurl, "_blank");
    } else {
      if (!file.documentbody || !file.mimetype || !file.filename) {
        toast.error(getLocalString(this.props.context!, LocalStrings.Toast.Message_Download_Error_Notes));
        return;
      }
      try {
        toast.loading(getLocalString(this.props.context!, LocalStrings.Toast.Message_Download_Prepare_Error_Notes));
        const blob = b64toBlob(file.documentbody, file.mimetype);
        const url = URL.createObjectURL(blob);
        const a = document.createElement("a");
        a.href = url;
        a.download = file.filename;
        document.body.appendChild(a);
        a.click();
        document.body.removeChild(a);
        URL.revokeObjectURL(url);
        toast.dismiss();
      } catch (e) {
        toast.dismiss();
        toast.error(getLocalString(this.props.context!, LocalStrings.Toast.Message_Download_Open_Error_SharePoint));
      }
    }
  }

  toggleSharePointDocLoc = async (
    _: React.MouseEvent<HTMLElement>,
    checked?: boolean
  ) => {
    const sharePointDocLoc = !!checked;
    const { selectedDocumentLocation, selectedDocumentLocationName } =
      this.state;
    this.setState(
      { sharePointDocLoc, selectedFiles: [], isLoading: true },
      async () => {
        this.loadExistingFiles();
      }
    );
  };

  toggleRememberLocation = async (
    _: React.MouseEvent<HTMLElement>,
    checked?: boolean
  ) => {
    const sharePointDocLoc = !!checked;
    const { selectedDocumentLocation, selectedDocumentLocationName } =
      this.state;
    this.setState({ userPreference: !!checked }, async () => {
      await this.saveUserSettings(
        sharePointDocLoc,
        selectedDocumentLocation,
        selectedDocumentLocationName
      );
    });
  };

  getFilteredAndSortedFiles() {
    const { files, notesSearchText } = this.state;
    return files.filter(
      (file) =>
        file.filename &&
        file.filename.toLowerCase().includes(notesSearchText.toLowerCase())
    );
  }

  getFilteredAndSortedSPFiles() {
    const { sharePointData, spSearchText } = this.state;
    return sharePointData
      .filter((spData) =>
        spData.fullname
          .toLowerCase()
          .includes(spSearchText.toLowerCase().trim())
      )
      .sort((a, b) => {
        if (a.filetype === "folder" && b.filetype !== "folder") {
          return -1;
        } else if (a.filetype !== "folder" && b.filetype === "folder") {
          return 1;
        }
        return a.fullname.toLowerCase().localeCompare(b.fullname.toLowerCase());
      });
  }

  async handleDropdownChange(
    event: React.FormEvent<HTMLDivElement>,
    option?: IDropdownOption
  ): Promise<void> {
    const userPreference = false;
    if (option) {
      const selectedDocumentLocation = option.key as string;
      const selectedDocumentLocationName = option.text as string;

      this.setState(
        {
          userPreference,
          selectedDocumentLocation,
          selectedDocumentLocationName,
        },
        async () => {
          await this.saveUserSettings(
            userPreference,
            selectedDocumentLocation,
            selectedDocumentLocationName
          );
          this.setState({ currentFolderPath: "" }, () => {
            this.loadExistingFiles();
          });
        }
      );
    }
  }

  async handleViewChange(
    _e: React.FormEvent<HTMLDivElement>,
    option?: IDropdownOption
  ) {
    if (!option) return;

    this.setState({ selectedViewId: option.key as string }, () =>
      this.loadExistingFiles()
    );
  }

  renderRibbon = () => {
    const { selectedFiles, sharePointDocLoc, isCollapsed, formType } =
      this.state;
    const formIsDisabled = this.props.isDisabled || formType !== 2;
    const commandButtons = (
      <Stack horizontal tokens={ribbonStackTokens}>
        {!sharePointDocLoc && selectedFiles.length < 2 && (
          <>
            <CommandButton
              text="Preview"
              iconProps={{ iconName: "View" }}
              onClick={() => this.performActionOnSelectedFiles("preview")}
              disabled={
                selectedFiles.length !== 1 ||
                !this.state.files.some(
                  (file) =>
                    selectedFiles.includes(file.noteId!) &&
                    (isImage(file.mimetype!) || isPDF(file.mimetype!))
                )
              }
              className="icon-button"
            />
            <CommandButton
              iconProps={{ iconName: "Edit" }}
              text="Rename"
              onClick={() => this.performActionOnSelectedFiles("edit")}
              disabled={
                selectedFiles.length !== 1 ||
                !this.state.files.some(
                  (file) =>
                    selectedFiles.includes(file.noteId!) &&
                    (isImage(file.mimetype!) || isPDF(file.mimetype!))
                )
              }
              className="icon-button"
            />
          </>
        )}
        <CommandButton
          text={`${getLocalString(this.props.context!, LocalStrings.Button.Label_Download)}`}
          iconProps={{ iconName: "Download" }}
          onClick={() => this.performActionOnSelectedFiles("download")}
          disabled={selectedFiles.length === 0 || formIsDisabled}
          className="icon-button"
        />
        <CommandButton
          iconProps={{ iconName: "Delete" }}
          text={`${getLocalString(this.props.context!, LocalStrings.Button.Label_Delete)}`}
          onClick={() => this.performActionOnSelectedFiles("delete")}
          disabled={selectedFiles.length === 0 || formIsDisabled}
          className="icon-button"
        />
        {sharePointDocLoc && this.state.isActivity && (
          <CommandButton
            iconProps={{ iconName: "Attach" }}
            text="Add to Activity"
            onClick={() => this.performActionOnSelectedFiles("addToActivityAttachment")}
            disabled={selectedFiles.length === 0 || formIsDisabled}
            className="icon-button"
          />
        )}
      </Stack>
    );

    const { context } = this.props;
    const { noteViews, selectedViewId, isViewsLoading } = this.state;

    let displayNotesViewsControlParameter =
      getControlValue(context!, "displayNotesViewsControl") === true;
    let enableNotesViewsControlParameter =
      getControlValue(context!, "enableNotesViewsControl") === true;

    // CUSTOM 2026-04-08: In SP-only mode, Notes view controls are never shown.
    const sharePointOnlyMode =
      getControlValue(context!, "sharePointOnlyMode") !== false;
    if (sharePointOnlyMode) {
      displayNotesViewsControlParameter = false;
      enableNotesViewsControlParameter = false;
    }

    const viewOptions: IDropdownOption[] = noteViews.map(v => ({
      key: v.savedqueryid,
      text: v.name,
    }));

    const menuItems: IContextualMenuItem[] = [
      {
        key: "download",
        text: `${getLocalString(this.props.context!, LocalStrings.Button.Label_Download)}`,
        iconProps: { iconName: "Download" },
        onClick: () => this.performActionOnSelectedFiles("download"),
        disabled: selectedFiles.length === 0 || formIsDisabled,
      },
      {
        key: "delete",
        text: `${getLocalString(this.props.context!, LocalStrings.Button.Label_Delete)}`,
        iconProps: { iconName: "Delete" },
        onClick: () => this.performActionOnSelectedFiles("delete"),
        disabled: selectedFiles.length === 0 || formIsDisabled,
      },
    ];

    if (!sharePointDocLoc) {
      menuItems.push(
        {
          key: "divider_1",
          itemType: ContextualMenuItemType.Divider,
        },
        {
          key: "preview",
          text: `${getLocalString(this.props.context!, LocalStrings.Button.Label_Preview)}`,
          iconProps: { iconName: "View" },
          onClick: () => this.performActionOnSelectedFiles("preview"),
          disabled:
            selectedFiles.length !== 1 ||
            !this.state.files.some(
              (file) =>
                selectedFiles.includes(file.noteId!) &&
                (isImage(file.mimetype!) || isPDF(file.mimetype!))
            ),
        },
        {
          key: "rename",
          text: `${getLocalString(this.props.context!, LocalStrings.Button.Label_Rename)}`,
          iconProps: { iconName: "Edit" },
          onClick: () => this.performActionOnSelectedFiles("edit"),
          disabled:
            selectedFiles.length === 0 ||
            selectedFiles.length >= 2 ||
            formIsDisabled,
        }
      );
    }
    return (
      <>
        <Stack
          horizontal
          styles={searchRibbonStyles}
          className="easeIn wideRibbon"
        >
          <Stack horizontal styles={ribbonStyles} className="easeIn">
            <Stack horizontal tokens={ribbonStackTokens}>
              <SearchBox
                placeholder={`${getLocalString(this.props.context!, LocalStrings.Input.Placeholder_Search)}`}
                onSearch={(newValue) => this.handleSearch(null, newValue)}
                onChange={(_, newValue) => this.handleSearch(null, newValue)}
                styles={{
                  root: {
                    width: 200,
                    border: "none",
                    boxShadow: "none",
                    position: "relative",
                    selectors: {
                      ":hover": {
                        border: "none",
                        boxShadow: "none",
                      },
                      ":focus": {
                        border: "none",
                        boxShadow: "none",
                      },
                      "::after": {
                        content: "none !important",
                      },
                    },
                  },
                  field: {
                    border: "none",
                    boxShadow: "none",
                  },
                  iconContainer: {
                    color: "rgb(17, 94, 163)",
                    selectors: {
                      ":hover": {
                        color: "rgb(17, 94, 163)",
                      },
                      ":focus": {
                        color: "rgb(17, 94, 163)",
                      },
                    },
                  },
                }}
              />
            </Stack>
            {selectedFiles.length > 0 && (
              <Stack horizontal tokens={ribbonStackTokens}>
                <>
                  {this.state.menuVisible && (
                    <ContextualMenu
                      items={menuItems}
                      target={this.state.target}
                      onDismiss={this.closeMenu}
                      isBeakVisible={true}
                    />
                  )}
                  {isCollapsed ? (
                    <CommandButton
                      iconProps={{ iconName: "More" }}
                      text={`${getLocalString(this.props.context!, LocalStrings.Button.Label_Actions)}`}
                      onClick={this.toggleMenu}
                    />
                  ) : (
                    commandButtons
                  )}
                </>
              </Stack>
            )}
            {sharePointDocLoc && (
              <Stack horizontal tokens={ribbonStackTokens} verticalAlign="center">
                <CommandButton
                  iconProps={{ iconName: "Add" }}
                  text={`${getLocalString(this.props.context!, LocalStrings.Button.Label_CreateFolder)}`}
                  onClick={this.toggleModal}
                  disabled={formIsDisabled}
                  className="icon-button"
                />
                {this.state.currentFolderPath && (
                  <CommandButton
                    iconProps={{ iconName: "Back" }}
                    text={`${getLocalString(this.props.context!, LocalStrings.Button.Label_Back)}`}
                    onClick={this.handleBackClick}
                    className="icon-button"
                  />
                )}
              </Stack>
            )}
            {!sharePointDocLoc && !sharePointOnlyMode && displayNotesViewsControlParameter && noteViews.length > 0 && (
              <Dropdown
                options={viewOptions}
                selectedKey={selectedViewId}
                onChange={(e, o) => this.handleViewChange(e, o)}
                disabled={!enableNotesViewsControlParameter || isViewsLoading}
                styles={dropdownStyles}
              />
            )}
          </Stack>
        </Stack>
      </>
    );
  };

  toggleModal = () => {
    this.setState((prevState) => ({ showModal: !prevState.showModal }));
  };

  handleFolderNameChange = (
    event: React.FormEvent<HTMLInputElement | HTMLTextAreaElement>,
    newValue?: string
  ) => {
    this.setState({ newFolderName: newValue || "" });
  };

  createSharePointFolder = async () => {
    const { newFolderName, selectedDocumentLocation, currentFolderPath, selectedDocumentLocationName } = this.state;
    if (!newFolderName.trim()) return;
    const defaultLocation = this.isDefaultLocation();
    const result = await createSharePointFolder(
      this.props.context!,
      newFolderName,
      selectedDocumentLocation,
      currentFolderPath,
      selectedDocumentLocationName,
      defaultLocation
    );
    if (result.success) {
      toast.success(`Folder "${newFolderName}" created.`);
      this.setState({ newFolderName: "", showModal: false }, () => {
        this.loadExistingFiles();
      });
    } else {
      toast.error(`Failed to create folder: ${result.message}`);
    }
  };

  handleFolderClick = (folderPath: string) => {
    this.setState(
      (prevState) => ({
        folderStack: [...prevState.folderStack, prevState.currentFolderPath],
        currentFolderPath: folderPath,
      }),
      () => {
        this.getSharePointData();
      }
    );
  };

  handleBackClick = () => {
    this.getSharePointData(true);
  };

  handleRemoveFolderClick = (folder: SharePointDocument) => {
    this.setState({
      isFolderDeletionDialogVisible: true,
      selectedFolderForDelete: folder,
    });
  };

  async saveUserSettings(
    sharePointEnabled: boolean,
    selectedDocumentLocation: string | null,
    selectedDocumentLocationName: string
  ) {
    const metadata = await getEntityMetadata(this.props.context!);
    const settings = {
      tableName: metadata?.schemaName,
      NotesOrSharePoint: sharePointEnabled,
      selectedDocumentLocation: selectedDocumentLocation,
      selectedDocumentLocationName: selectedDocumentLocationName,
    };
    localStorage.setItem("userSettings", JSON.stringify(settings));
  }

  performActionOnSelectedFiles = async (action: FileAction) => {
    const { selectedFiles } = this.state;

    selectedFiles.forEach((fileId) => {
      const file = this.state.sharePointDocLoc
        ? this.state.sharePointData.find(
          (doc) => doc.sharepointdocumentid === fileId
        )
        : this.state.files.find((file) => file.noteId === fileId);

      if (!file) {
        toast.error(`File with ID ${fileId} not found.`);
        return;
      }

      if (this.state.sharePointDocLoc && "sharepointdocumentid" in file) {
        switch (action) {
          case "addToActivityAttachment":
            this.addFileAttachmentToActivity(file);
            break;
          case "download":
            this.downloadFile(file);
            break;
          case "delete":
            this.removeFile(file.sharepointdocumentid);
            break;
          default:
            toast.error(
              `Action ${action} is not supported for SharePoint documents.`
            );
        }
      } else if (!this.state.sharePointDocLoc && "noteId" in file) {
        switch (action) {
          case "download":
            this.downloadFile(file);
            break;
          case "delete":
            this.removeFile(file.noteId);
            break;
          case "edit":
            this.toggleEditModal(file.noteId);
            break;
          case "preview":
            this.openDialog({
              filename: file.filename,
              documentbody: file.documentbody!,
              mimetype: file.mimetype!,
            });
            break;
          default:
            toast.error(`Action ${action} is not supported for notes.`);
        }
      }
    });

    this.setState({ selectedFiles: [] });
  };

  openDialog(file: PreviewFile) {
    const isImageType = isImage(file.mimetype);
    if (isImageType) {
      this.setState({
        previewFile: {
          ...file,
          documentbody: createDataUri(file.mimetype, file.documentbody),
        },
        xlsxContent: "",
        xlsxData: null,
      });
    } else {
      this.setState({
        isDialogOpen: true,
        previewFile: {
          ...file,
          documentbody: createDataUri(file.mimetype, file.documentbody),
        },
        xlsxContent: "",
        xlsxData: null,
      });
    }
  }

  closeDialog() {
    this.setState({ isDialogOpen: false, previewFile: null });
  }

  handleSearch = (event: any, newValue?: string) => {
    const searchText = newValue || "";

    if (this.state.sharePointDocLoc) {
      this.setState({ spSearchText: searchText });
    } else {
      this.setState({ notesSearchText: searchText });
    }
  };

  toggleSortOrder = () => {
    this.setState((prevState) => ({ sortAsc: !prevState.sortAsc }));
  };

  toggleTooltip() {
    this.setState((prevState) => ({ showTooltip: !prevState.showTooltip }));
  }

  toggleEditModal(noteId?: string) {
    this.setState({ editingFileId: noteId });
  }

  async saveChanges(noteId: string, filename: string) {
    const result = await updateRelatedNote(
      this.props.context!,
      noteId,
      filename
    );
    if (result.success) {
      toast.success("File renamed.");
      this.toggleEditModal();
      this.loadExistingFiles();
    } else {
      toast.error(`Failed to rename: ${result.message}`);
    }
  }

  renderFileList() {
    const { sharePointDocLoc } = this.state;
    if (sharePointDocLoc) {
      return this.renderSharePointFiles();
    } else {
      return this.renderNoteFiles();
    }
  }

  renderNoteFiles() {
    const files = this.getFilteredAndSortedFiles();
    return (
      <div className="file-list">
        {files.map((file) => (
          <div
            key={file.noteId}
            className={`file-box ${this.state.selectedFiles.includes(file.noteId!) ? "selected" : ""}`}
            onClick={() => this.toggleFileSelection(file.noteId!)}
          >
            <Icon
              {...getFileTypeIconProps({
                extension: this.getFileExtension(file.filename),
                size: 32,
              })}
            />
            <p title={file.filename}>{file.filename}</p>
            <p className="file-size">{this.formatFileSize(file.filesize)}</p>
          </div>
        ))}
      </div>
    );
  }

  renderSharePointFiles() {
    const files = this.getFilteredAndSortedSPFiles();
    return (
      <div className="file-list">
        {files.map((file) => (
          <div
            key={file.sharepointdocumentid}
            className={`file-box ${this.state.selectedFiles.includes(file.sharepointdocumentid) ? "selected" : ""}`}
            onClick={() =>
              file.filetype === "folder"
                ? this.handleFolderClick(file.locationurl)
                : this.toggleFileSelection(file.sharepointdocumentid)
            }
          >
            {file.filetype === "folder" ? (
              <>
                <div style={{ position: "relative", display: "inline-block" }}>
                  <FolderIcon />
                  <div
                    className="remove-button"
                    onClick={(e) => {
                      e.stopPropagation();
                      this.handleRemoveFolderClick(file);
                    }}
                  >
                    <IconButton
                      iconProps={{ iconName: "Cancel" }}
                      styles={{
                        root: {
                          width: 16,
                          height: 16,
                          color: "gray",
                        },
                        icon: { fontSize: 10 },
                      }}
                    />
                  </div>
                </div>
                <p title={file.fullname}>{file.fullname}</p>
              </>
            ) : (
              <>
                <Icon
                  {...getFileTypeIconProps({
                    extension: this.getFileExtension(file.fullname),
                    size: 32,
                  })}
                />
                <p title={file.fullname}>{file.fullname}</p>
                <p className="file-size">{this.formatFileSize(file.filesize)}</p>
              </>
            )}
          </div>
        ))}
      </div>
    );
  }

  toggleFileSelection(fileId: string) {
    this.setState((prevState) => {
      const selectedFiles = prevState.selectedFiles.includes(fileId)
        ? prevState.selectedFiles.filter((id) => id !== fileId)
        : [...prevState.selectedFiles, fileId];
      return { selectedFiles };
    });
  }

  formatFileSize(size: number): string {
    if (size < 1024) return `${size} B`;
    if (size < 1024 * 1024) return `${(size / 1024).toFixed(1)} KB`;
    return `${(size / (1024 * 1024)).toFixed(1)} MB`;
  }

  render() {
    const {
      editingFileId,
      isLoading,
      sharePointDocLoc,
      showTooltip,
      selectedDocumentLocation,
      documentLocations,
      showModal,
      newFolderName,
      isDialogOpen,
      previewFile,
      sharePointData,
      sharePointEnabled,
      currentFolderPath,
      isCalloutVisible,
      isCreateLocationDialogVisible,
      userPreference,
      createLocationDisplayName,
      selectedCreateLocation,
      createLocationFolderName,
      isSaveButtonEnabled,
      sharePointEnabledParameter,
      formType,
      noteViews,
      selectedViewId
    } = this.state;

    const formIsDisabled = formType !== 2;

    // ============================================================
    // CUSTOM 2026-04-08: Resolve sharePointOnlyMode once for render.
    // All UI gating in the render tree uses this local const.
    // ============================================================
    const { context, isDisabled } = this.props;
    const sharePointOnlyMode =
      getControlValue(context!, "sharePointOnlyMode") !== false;

    const dropdownOptions: IDropdownOption[] = documentLocations.map(
      (location) => ({
        key: location.sharepointdocumentlocationid,
        text: location.name,
      })
    );
    const viewOptions: IDropdownOption[] = noteViews.map(v => ({
      key: v.savedqueryid,
      text: v.name,
    }));
    const comboBoxOptions: IDropdownOption[] = documentLocations
      .filter((location) => location.parentsiteid)
      .map((location) => ({
        key: location.parentsiteid!,
        text: location.name,
      }));
    const files = this.getFilteredAndSortedFiles();
    const addIcon: IIconProps = { iconName: "Add" };
    const splitButtonStyles: IButtonStyles = {
      splitButtonMenuButton: {
        backgroundColor: "white",
        width: 28,
        border: "none",
      },
      splitButtonMenuIcon: { fontSize: "7px" },
      splitButtonDivider: {
        backgroundColor: "#c8c8c8",
        width: 1,
        right: 26,
        position: "absolute",
        top: 4,
        bottom: 4,
      },
      splitButtonContainer: {
        selectors: {
          [HighContrastSelector]: { border: "none" },
        },
      },
    };
    const menuProps: IContextualMenuProps = {
      items: [
        {
          key: "settings",
          iconProps: { iconName: "Settings" },
          text: `${getLocalString(this.props.context!, LocalStrings.Button.Label_Settings)}`,
          onClick: this.onGearIconClick,
        },
      ],
    };
    const isEmpty = sharePointDocLoc
      ? sharePointData.length === 0
      : files.length === 0;
    const entityIdExists = (this.props.context as any).page.entityId;
    const editingFile = files.find((file) => file.noteId === editingFileId);
    const docLoctooltip = (
      <TooltipHost
        content={showTooltip ? SharePointDocLocTooltip : undefined}
        id={tooltipId}
        closeDelay={500}
      >
        <DefaultButton
          aria-label="more info"
          aria-describedby={showTooltip ? tooltipId : undefined}
          onClick={this.toggleTooltip}
          styles={buttonStyles}
          iconProps={{ iconName: "Info" }}
        />
      </TooltipHost>
    );

    if (!entityIdExists) {
      return (
        <div className="record-not-created-message">
          {getLocalString(this.props.context!, LocalStrings.Input.Placeholder_Record_Not_Created)}
        </div>
      );
    }

    // Resolve SP dropdown display parameters, overridden by sharePointOnlyMode
    let displaySPDropdownControlParameter =
      getControlValue(context!, "displaySharePointDocumentLocationsControl") === true;
    let enableSharePointDocumentLocationsControlParameter =
      getControlValue(context!, "enableSharePointDocumentLocationsControl") === true;

    // CUSTOM 2026-04-08: In SP-only mode, force the document-location dropdown
    // to hidden regardless of the form property setting.
    if (sharePointOnlyMode) {
      displaySPDropdownControlParameter = false;
      enableSharePointDocumentLocationsControlParameter = false;
    }

    const base64 = previewFile?.documentbody.replace(
      /^data:application\/pdf;base64,/i,
      ""
    );
    const pdfDataUri = `data:application/pdf;base64,${base64}`;

    return (
      <>
        <Toaster position="top-right" reverseOrder={false} />
        {isImage(previewFile?.mimetype ?? "") && previewFile && (
          <ReactViewer
            visible
            onClose={this.closeDialog}
            container={this.viewerHost.current!}
            images={[{
              src: previewFile.documentbody.replace(
                /(data:image\/[a-zA-Z]+;base64,).*?\1/, '$1'),
              alt: previewFile.filename
            }]}
            rotatable scalable zoomable
            noNavbar={true}
          />
        )}

        {isDialogOpen && previewFile && (
          <Modal
            isOpen={isDialogOpen}
            onDismiss={this.closeDialog}
            isBlocking={false}
            containerClassName="ms-modalExample-container"
            styles={{
              main: {
                width: "100%",
                height: "100%",
                "@media(max-width: 767px)": { width: "100%" },
              },
              scrollableContent: {
                height: "100%",
                overflow: "hidden",
              },
            }}
          >
            <div className={"scrollable-area"}>
              <div ref={this.viewerHost} style={{ position: 'relative', height: 0 }} />
              {isPDF(previewFile.mimetype) && (
                previewFile.documentbody.length > 1_000_000 ? (
                  <iframe
                    src={pdfDataUri}
                    style={{ width: "100%", height: "100%", border: "none" }}
                    title={previewFile.filename}
                  />
                ) : (
                  <iframe
                    src={pdfDataUri}
                    style={{ width: "100%", height: "100%", border: "none" }}
                    title={previewFile.filename}
                  />
                )
              )}
            </div>
          </Modal>
        )}

        {editingFile && (
          <Dialog
            hidden={!editingFileId}
            onDismiss={() => this.toggleEditModal()}
            dialogContentProps={{
              type: DialogType.normal,
              title: `${getLocalString(this.props.context!, LocalStrings.Dialog.EditFile.Title)}`,
            }}
          >
            <TextField
              value={editingFile.filename}
              onChange={(_, newValue) => {
                const updatedFiles = this.state.files.map((file) =>
                  file.noteId === editingFileId
                    ? { ...file, filename: newValue || "" }
                    : file
                );
                this.setState({ files: updatedFiles });
              }}
            />
            <DialogFooter>
              <PrimaryButton
                onClick={() => {
                  this.saveChanges(editingFile.noteId!, editingFile.filename!);
                }}
                text={`${getLocalString(this.props.context!, LocalStrings.Dialog.EditFile.Button_Save)}`}
              />
              <DefaultButton
                onClick={() => this.toggleEditModal()}
                text={`${getLocalString(this.props.context!, LocalStrings.Dialog.EditFile.Button_Cancel)}`}
              />
            </DialogFooter>
          </Dialog>
        )}

        <Dialog
          hidden={!showModal}
          onDismiss={this.toggleModal}
          dialogContentProps={{
            type: DialogType.normal,
            title: `${getLocalString(this.props.context!, LocalStrings.Dialog.CreateFolder.Title)}`,
            subText: `${getLocalString(this.props.context!, LocalStrings.Dialog.CreateFolder.Description)}`,
          }}
          modalProps={{
            isBlocking: false,
            styles: { main: { maxWidth: 450 } },
          }}
        >
          <TextField
            value={newFolderName}
            onChange={this.handleFolderNameChange}
            placeholder={`${getLocalString(this.props.context!, LocalStrings.Dialog.CreateFolder.Input_Placeholder)}`}
          />
          <DialogFooter>
            <PrimaryButton
              onClick={this.createSharePointFolder}
              disabled={!newFolderName.trim()}
              text={`${getLocalString(this.props.context!, LocalStrings.Dialog.CreateFolder.Button_Create)}`}
            />
            <DefaultButton onClick={this.toggleModal} text={`${getLocalString(this.props.context!, LocalStrings.Dialog.CreateFolder.Button_Cancel)}`} />
          </DialogFooter>
        </Dialog>

        <Dialog
          hidden={!this.state.isFolderDeletionDialogVisible}
          onDismiss={() => this.setState({ isFolderDeletionDialogVisible: false, selectedFolderForDelete: null })}
          dialogContentProps={{
            type: DialogType.normal,
            title: `${getLocalString(this.props.context!, LocalStrings.Button.Label_Remove_Folder)}`,
            subText: `Are you sure you want to delete "${this.state.selectedFolderForDelete?.fullname}"?`,
          }}
        >
          <DialogFooter>
            <PrimaryButton
              onClick={async () => {
                if (this.state.selectedFolderForDelete) {
                  await this.removeFile(this.state.selectedFolderForDelete.sharepointdocumentid);
                  this.setState({ isFolderDeletionDialogVisible: false, selectedFolderForDelete: null });
                }
              }}
              text={`${getLocalString(this.props.context!, LocalStrings.Button.Label_Remove)}`}
            />
            <DefaultButton
              onClick={() => this.setState({ isFolderDeletionDialogVisible: false, selectedFolderForDelete: null })}
              text={`${getLocalString(this.props.context!, LocalStrings.Dialog.CreateFolder.Button_Cancel)}`}
            />
          </DialogFooter>
        </Dialog>

        <div className="ribbon-dropzone-wrapper">
          {this.renderRibbon()}

          <Dropzone onDrop={this.handleDrop} disabled={formIsDisabled || isDisabled}>
            {({ getRootProps, getInputProps }) => (
              <div className="dropzone-wrapper">
                <div
                  {...getRootProps()}
                  className={`dropzone ${isEmpty ? "empty" : ""}`}
                  style={{ cursor: (formIsDisabled || isDisabled) ? "default" : "pointer" }}
                >
                  {isLoading ? (
                    <div className="spinner-box">
                      <Spinner size={SpinnerSize.large} />
                    </div>
                  ) : (
                    <>
                      <input {...getInputProps()} />
                      <div style={{ width: "100%" }}>
                        <div style={{ width: "100%" }}>
                          {isEmpty ? (
                            <p>{`${getLocalString(this.props.context!, LocalStrings.Input.Placeholder_Dropzone)}`}</p>
                          ) : (
                            this.renderFileList()
                          )}
                        </div>
                      </div>
                    </>
                  )}
                </div>
              </div>
            )}
          </Dropzone>

          {/*
            ============================================================
            CUSTOM 2026-04-08: Bottom bar (toggle + SP dropdown + gear)
            In sharePointOnlyMode ALL three UI elements are hidden:
              - SharePoint/Notes toggle         → hidden
              - SP document-location dropdown   → hidden (displaySPDropdownControlParameter forced false above)
              - "Remember Location" gear callout → hidden

            In classic mode (sharePointOnlyMode = false) the original
            behaviour is fully preserved.
            ============================================================
          */}
          {!sharePointOnlyMode && (
            <Stack
              horizontal
              style={{
                width: "100%",
                justifyContent: !sharePointEnabledParameter ? "flex-end" : "space-between",
                alignItems: "center",
                flexWrap: "wrap",
                marginTop: "8px"
              }}
            >
              {/* SharePoint / Notes toggle */}
              {sharePointEnabledParameter && (
                <Toggle
                  label={<div>SharePoint Documents {docLoctooltip}</div>}
                  inlineLabel
                  onText={`${getLocalString(this.props.context!, LocalStrings.Toggle.Value_On)}`}
                  offText={`${getLocalString(this.props.context!, LocalStrings.Toggle.Value_Off)}`}
                  checked={sharePointDocLoc}
                  onChange={this.toggleSharePointDocLoc}
                  disabled={sharePointEnabled}
                  styles={{ root: { marginBottom: "0px" } }}
                />
              )}

              <Stack horizontal verticalAlign="center" tokens={{ childrenGap: 5 }}>
                {/* "Remember Location" settings callout */}
                {isCalloutVisible && (
                  <Callout
                    target={this.targetRef.current}
                    onDismiss={this.onCalloutDismiss}
                    directionalHint={DirectionalHint.bottomAutoEdge}
                    setInitialFocus
                    role="dialog"
                  >
                    <Stack tokens={{ childrenGap: 10 }} padding={20}>
                      <Stack horizontal verticalAlign="center" tokens={{ childrenGap: 10 }}>
                        <Label>Remember Location?</Label>
                        <Toggle
                          onChange={this.toggleRememberLocation}
                          styles={{ root: { marginBottom: "0px" } }}
                          checked={userPreference}
                        />
                      </Stack>
                    </Stack>
                  </Callout>
                )}

                {/* SP document-location dropdown (classic mode) */}
                {sharePointDocLoc && displaySPDropdownControlParameter && sharePointEnabledParameter ? (
                  <Dropdown
                    options={dropdownOptions}
                    selectedKey={selectedDocumentLocation}
                    onChange={(e, o) => this.handleDropdownChange(e, o)}
                    disabled={!enableSharePointDocumentLocationsControlParameter}
                    styles={dropdownStyles}
                  />
                ) : null}

                {/* Gear / settings button */}
                {sharePointDocLoc && sharePointEnabledParameter && (
                  <div ref={this.targetRef}>
                    <IconButton
                      menuProps={menuProps}
                      iconProps={{ iconName: "Settings" }}
                      onRenderMenuIcon={() => null}
                      onClick={this.onGearIconClick}
                    />
                  </div>
                )}
              </Stack>
            </Stack>
          )}
          {/* END CUSTOM: bottom bar hidden in sharePointOnlyMode */}

        </div>

        {/* Create Location dialog – unchanged, available in both modes */}
        {isCreateLocationDialogVisible && (
          <Dialog
            hidden={!isCreateLocationDialogVisible}
            onDismiss={() => this.setState({ isCreateLocationDialogVisible: false })}
            dialogContentProps={{
              type: DialogType.normal,
              title: "Create SharePoint Location",
            }}
          >
            <TextField
              label="Display Name"
              value={createLocationDisplayName}
              onChange={(_, v) => this.setState({ createLocationDisplayName: v || "" })}
            />
            <ComboBox
              label="Parent Site"
              options={comboBoxOptions}
              selectedKey={selectedCreateLocation}
              onChange={(_: any, option?: IComboBoxOption) => {
                if (option) this.setState({ selectedCreateLocation: option.key as string });
              }}
            />
            <TextField
              label="Folder Name"
              value={createLocationFolderName}
              onChange={(_, v) => this.setState({ createLocationFolderName: v || "", isSaveButtonEnabled: !!v })}
            />
            <DialogFooter>
              <PrimaryButton
                disabled={!isSaveButtonEnabled}
                onClick={async () => {
                  const res = await createSharePointLocation(
                    this.props.context!,
                    createLocationDisplayName,
                    selectedCreateLocation,
                    createLocationFolderName
                  );
                  if (res.success) {
                    toast.success("Location created.");
                    this.setState({ isCreateLocationDialogVisible: false }, () =>
                      this.loadExistingFiles()
                    );
                  } else {
                    toast.error(`Failed: ${res.message}`);
                  }
                }}
                text="Save"
              />
              <DefaultButton
                onClick={() => this.setState({ isCreateLocationDialogVisible: false })}
                text="Cancel"
              />
            </DialogFooter>
          </Dialog>
        )}
      </>
    );
  }
}
