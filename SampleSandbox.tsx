import * as React from 'react';
import { useCallback, useMemo, useState } from 'react';
import styles from './SampleSandbox.module.scss';
import type { ISampleSandboxProps } from './ISampleSandboxProps';
import {
  ComboBox,
  DefaultButton,
  Icon,
  IComboBoxOption,
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

interface ITipForm {
  productToolName: string;
  tipName: string;
  tipDescription: string;
}

interface ISPPeoplePickerUserEntity {
  Description: string;
  DisplayText: string;
  EntityData?: {
    Email?: string;
  };
  Key: string;
}

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
  const [authors, setAuthors] = useState<IPersonaProps[]>(defaultAuthor);
  const [form, setForm] = useState<ITipForm>({
    productToolName: '',
    tipName: '',
    tipDescription: ''
  });

  const updateField = useCallback((field: keyof ITipForm, value: string) => {
    setForm((prev) => ({ ...prev, [field]: value }));
  }, []);

  const resetForm = useCallback((): void => {
    setForm({
      productToolName: '',
      tipName: '',
      tipDescription: ''
    });
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

  const handleSubmit = useCallback((): void => {
    if (
      !form.productToolName.trim() ||
      !form.tipName.trim() ||
      !form.tipDescription.trim() ||
      authors.length === 0
    ) {
      return;
    }

    console.log('Tip submitted:', {
      ...form,
      author: authors[0]
    });

    setSubmitted(true);
    setIsModalOpen(false);
    resetForm();
  }, [authors, form, resetForm]);

  const isFormValid =
    form.productToolName.trim().length > 0 &&
    form.tipName.trim().length > 0 &&
    form.tipDescription.trim().length > 0 &&
    authors.length > 0;

  return (
    <section className={styles.sampleSandbox}>
      <div className={styles.content}>
        <h2 className={styles.title}>Product Tips</h2>
        <p className={styles.description}>
          Share helpful tips about products and tools with your team.
        </p>

        {submitted && (
          <MessageBar messageBarType={MessageBarType.success} isMultiline={false}>
            Your tip was submitted successfully.
          </MessageBar>
        )}

        <PrimaryButton text="Submit a Tip" onClick={openModal} />

        <Modal
          isOpen={isModalOpen}
          onDismiss={closeModal}
          isBlocking={false}
          containerClassName={styles.modalContainer}
        >
          <div className={styles.modalContent}>
            <div className={styles.modalTitleRow}>
              <Icon iconName="Lightbulb" className={styles.modalTitleIcon} />
              <h2 className={styles.modalTitle}>Submit a Product Tip</h2>
            </div>

            <Stack tokens={stackTokens} className={styles.roundedFields}>
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

              <TextField
                label="Tip Name"
                required
                value={form.tipName}
                onChange={(_, value) => updateField('tipName', value || '')}
                placeholder="Short title for your tip"
              />

              <TextField
                label="Tip Description/Steps"
                required
                multiline
                rows={5}
                value={form.tipDescription}
                onChange={(_, value) => updateField('tipDescription', value || '')}
                placeholder="Describe the tip or list the steps"
              />

              <div>
                <Label required>Author of the Tip</Label>
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
