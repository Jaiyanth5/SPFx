import * as React from 'react';
import { useCallback, useMemo, useState } from 'react';
import styles from './SampleSandbox.module.scss';
import type { ISampleSandboxProps } from './ISampleSandboxProps';
import {
  ComboBox,
  DatePicker,
  DefaultButton,
  Dropdown,
  Icon,
  IComboBoxOption,
  IDropdownOption,
  IStackTokens,
  Label,
  MessageBar,
  MessageBarType,
  Modal,
  NormalPeoplePicker,
  PrimaryButton,
  Stack,
  TextField
} from '@fluentui/react';
import type { IBasePickerSuggestionsProps } from '@fluentui/react/lib/Pickers';
import type { IPersonaProps } from '@fluentui/react/lib/Persona';
import { initializeIcons } from '@fluentui/react/lib/Icons';
import { SPHttpClient } from '@microsoft/sp-http';

initializeIcons();

type EntryType = 'productTip' | 'incident';

interface IFormState {
  entryType: EntryType;
  productToolName: string;
  title: string;
  description: string;
  buImpact: string;
}

interface ISPPeoplePickerUserEntity {
  Description: string;
  DisplayText: string;
  EntityData?: {
    Email?: string;
  };
  Key: string;
}

const typeOptions: IDropdownOption[] = [
  { key: 'productTip', text: 'Product Tip' },
  { key: 'incident', text: 'Incident' }
];

const productOptions: IComboBoxOption[] = [
  { key: 'teams', text: 'Microsoft Teams' },
  { key: 'sharepoint', text: 'SharePoint' },
  { key: 'onedrive', text: 'OneDrive' },
  { key: 'outlook', text: 'Outlook' },
  { key: 'power-automate', text: 'Power Automate' },
  { key: 'power-apps', text: 'Power Apps' },
  { key: 'power-bi', text: 'Power BI' },
  { key: 'viva', text: 'Microsoft Viva' },
  { key: 'graph', text: 'Microsoft Graph' },
  { key: 'azure-devops', text: 'Azure DevOps' },
  { key: 'word', text: 'Word' },
  { key: 'excel', text: 'Excel' },
  { key: 'powerpoint', text: 'PowerPoint' },
  { key: 'loop', text: 'Loop' },
  { key: 'planner', text: 'Planner' },
  { key: 'forms', text: 'Forms' }
];

const stackTokens: IStackTokens = { childrenGap: 15 };
const buttonStackTokens: IStackTokens = { childrenGap: 8 };

const pickerSuggestionsProps: IBasePickerSuggestionsProps = {
  suggestionsHeaderText: 'Suggested People',
  noResultsFoundText: 'No results found',
  loadingText: 'Loading'
};

const emptyForm = (): IFormState => ({
  entryType: 'productTip',
  productToolName: '',
  title: '',
  description: '',
  buImpact: ''
});

const searchPeople = async (
  filterText: string,
  props: ISampleSandboxProps
): Promise<IPersonaProps[]> => {
  if (!filterText) {
    return [];
  }

  const endpoint =
    `${props.context.pageContext.web.absoluteUrl}` +
    '/_api/SP.UI.ApplicationPages.ClientPeoplePickerWebServiceInterface.clientPeoplePickerSearchUser';

  const response = await props.context.spHttpClient.post(
    endpoint,
    SPHttpClient.configurations.v1,
    {
      headers: {
        Accept: 'application/json',
        'Content-Type': 'application/json'
      },
      body: JSON.stringify({
        queryParams: {
          AllowEmailAddresses: true,
          AllowMultipleEntities: false,
          AllUrlZones: false,
          MaximumEntitySuggestions: 20,
          PrincipalSource: 15,
          PrincipalType: 1,
          QueryString: filterText
        }
      })
    }
  );

  if (!response.ok) {
    return [];
  }

  const json: { value: string } = await response.json();
  const users: ISPPeoplePickerUserEntity[] = JSON.parse(json.value);

  return users.map((user) => ({
    text: user.DisplayText,
    secondaryText: user.EntityData?.Email || user.Description,
    data: { userId: user.Key }
  }));
};

const SampleSandbox: React.FC<ISampleSandboxProps> = (props) => {
  const { userDisplayName, userEmail, userLoginName, context } = props;

  const defaultAuthor = useMemo((): IPersonaProps[] => {
    if (!userDisplayName) {
      return [];
    }

    return [{
      text: userDisplayName,
      secondaryText: userEmail,
      imageUrl: `${context.pageContext.web.absoluteUrl}/_layouts/15/userphoto.aspx?size=S&username=${encodeURIComponent(userLoginName)}`
    }];
  }, [context.pageContext.web.absoluteUrl, userDisplayName, userEmail, userLoginName]);

  const [isModalOpen, setIsModalOpen] = useState<boolean>(false);
  const [submitted, setSubmitted] = useState<boolean>(false);
  const [submittedType, setSubmittedType] = useState<EntryType>('productTip');
  const [authors, setAuthors] = useState<IPersonaProps[]>(defaultAuthor);
  const [resolvedDate, setResolvedDate] = useState<Date | undefined>(undefined);
  const [form, setForm] = useState<IFormState>(emptyForm);

  const isIncident = form.entryType === 'incident';

  const fieldLabels = useMemo(() => ({
    modalTitle: isIncident ? 'Submit an Incident' : 'Submit a Product Tip',
    modalIcon: isIncident ? 'Warning' : 'Lightbulb',
    title: isIncident ? 'Incident Title' : 'Product Tip Title',
    description: isIncident ? 'Incident Description' : 'Tip Description',
    author: isIncident ? 'Incident Reporter' : 'Author of the Tip',
    titlePlaceholder: isIncident ? 'Short title for the incident' : 'Short title for your tip',
    descriptionPlaceholder: isIncident
      ? 'Describe what happened and how it was resolved'
      : 'Describe the tip or list the steps'
  }), [isIncident]);

  const updateField = useCallback((field: keyof IFormState, value: string) => {
    setForm((prev) => ({ ...prev, [field]: value }));
  }, []);

  const resetForm = useCallback((): void => {
    setForm(emptyForm());
    setResolvedDate(undefined);
    setAuthors(defaultAuthor);
  }, [defaultAuthor]);

  const openModal = useCallback((): void => {
    setSubmitted(false);
    setIsModalOpen(true);
  }, []);

  const closeModal = useCallback((): void => {
    setIsModalOpen(false);
  }, []);

  const onResolveSuggestions = useCallback(
    (filterText: string): Promise<IPersonaProps[]> => searchPeople(filterText, props),
    [props]
  );

  const handleTypeChange = useCallback((_: React.FormEvent<HTMLDivElement>, option?: IDropdownOption) => {
    const entryType = (option?.key as EntryType) || 'productTip';
    setForm((prev) => ({
      ...prev,
      entryType,
      buImpact: entryType === 'incident' ? prev.buImpact : ''
    }));

    if (entryType === 'productTip') {
      setResolvedDate(undefined);
    }
  }, []);

  const handleSubmit = useCallback((): void => {
    const baseValid =
      form.productToolName.trim().length > 0 &&
      form.title.trim().length > 0 &&
      form.description.trim().length > 0 &&
      authors.length > 0;

    const incidentValid = !isIncident || (!!resolvedDate && form.buImpact.trim().length > 0);

    if (!baseValid || !incidentValid) {
      return;
    }

    console.log('Entry submitted:', {
      ...form,
      resolvedDate: isIncident ? resolvedDate : undefined,
      author: authors[0]
    });

    setSubmittedType(form.entryType);
    setSubmitted(true);
    setIsModalOpen(false);
    resetForm();
  }, [authors, form, isIncident, resolvedDate, resetForm]);

  const isFormValid =
    form.productToolName.trim().length > 0 &&
    form.title.trim().length > 0 &&
    form.description.trim().length > 0 &&
    authors.length > 0 &&
    (!isIncident || (!!resolvedDate && form.buImpact.trim().length > 0));

  return (
    <section className={styles.sampleSandbox}>
      <div className={styles.content}>
        <h2 className={styles.title}>Product Tips & Incidents</h2>
        <p className={styles.description}>
          Share helpful tips or report incidents about products and tools with your team.
        </p>

        {submitted && (
          <MessageBar messageBarType={MessageBarType.success} isMultiline={false}>
            {submittedType === 'incident'
              ? 'Your incident was submitted successfully.'
              : 'Your tip was submitted successfully.'}
          </MessageBar>
        )}

        <PrimaryButton text="Submit Entry" onClick={openModal} />

        <Modal
          isOpen={isModalOpen}
          onDismiss={closeModal}
          isBlocking={false}
          containerClassName={styles.modalContainer}
        >
          <div className={styles.modalContent}>
            <div className={styles.modalTitleRow}>
              <Icon iconName={fieldLabels.modalIcon} className={styles.modalTitleIcon} />
              <h2 className={styles.modalTitle}>{fieldLabels.modalTitle}</h2>
            </div>

            <Stack tokens={stackTokens} className={styles.roundedFields}>
              <Dropdown
                label="Type"
                required
                options={typeOptions}
                selectedKey={form.entryType}
                onChange={handleTypeChange}
              />

              <div className={styles.formRow}>
                <div className={styles.formField}>
                  <ComboBox
                    label="Product/Tool Name"
                    required
                    options={productOptions}
                    selectedKey={productOptions.find((option) => option.text === form.productToolName)?.key}
                    text={form.productToolName}
                    allowFreeInput
                    autoComplete="on"
                    useComboBoxAsMenuWidth
                    placeholder="Search or select a product/tool"
                    onChange={(_, option, __, value) => {
                      updateField('productToolName', option?.text || value || '');
                    }}
                  />
                </div>

                <div className={styles.formField}>
                  <TextField
                    label={fieldLabels.title}
                    required
                    value={form.title}
                    onChange={(_, value) => updateField('title', value || '')}
                    placeholder={fieldLabels.titlePlaceholder}
                  />
                </div>
              </div>

              <TextField
                label={fieldLabels.description}
                required
                multiline
                rows={5}
                value={form.description}
                onChange={(_, value) => updateField('description', value || '')}
                placeholder={fieldLabels.descriptionPlaceholder}
              />

              {isIncident && (
                <div className={styles.formRow}>
                  <div className={styles.formField}>
                    <DatePicker
                      label="Resolved Date"
                      isRequired
                      value={resolvedDate}
                      onSelectDate={(date) => setResolvedDate(date || undefined)}
                      placeholder="Select a date"
                    />
                  </div>

                  <div className={styles.formField}>
                    <TextField
                      label="BU Impact"
                      required
                      value={form.buImpact}
                      onChange={(_, value) => updateField('buImpact', value || '')}
                      placeholder="Business unit(s) affected"
                    />
                  </div>
                </div>
              )}

              <div>
                <Label required>{fieldLabels.author}</Label>
                <NormalPeoplePicker
                  className={styles.peoplePicker}
                  onResolveSuggestions={onResolveSuggestions}
                  getTextFromItem={(persona) => persona.text || ''}
                  pickerSuggestionsProps={pickerSuggestionsProps}
                  selectedItems={authors}
                  onChange={(items) => setAuthors(items || [])}
                  itemLimit={1}
                  inputProps={{
                    placeholder: 'Search for a person'
                  }}
                />
              </div>

              <Stack horizontal tokens={buttonStackTokens}>
                <PrimaryButton text="Submit" onClick={handleSubmit} disabled={!isFormValid} />
                <DefaultButton text="Cancel" onClick={closeModal} />
              </Stack>
            </Stack>
          </div>
        </Modal>
      </div>
    </section>
  );
};

export default SampleSandbox;
