import {
    OpenAIModel,
    PromptManager,
    ActionPlanner,
    Application,
    TurnState,
    TeamsAdapter,
    AI,
    Message
} from '@microsoft/teams-ai';
import {
    CardFactory,
    ConfigurationServiceClientCredentialFactory,
    MemoryStorage,
    MessageFactory,
    TurnContext
} from 'botbuilder';
import path from 'path';
import debug from 'debug';
import { createResponseCard } from './card';

const error = debug('azureopenai:app:error');
error.log = console.log.bind(console);

interface ConversationState {}
type ApplicationTurnState = TurnState<ConversationState>;

const deployment = process.env.AZURE_OPENAI_DEPLOYMENT || 'gpt-35-turbo-16k';

if (
    !process.env.AZURE_OPENAI_KEY ||
    !process.env.AZURE_OPENAI_ENDPOINT ||
    !process.env.AZURE_SEARCH_ENDPOINT ||
    !process.env.AZURE_SEARCH_KEY
) {
    throw new Error(
        'Missing environment variables - please check that AZURE_OPENAI_KEY, AZURE_OPENAI_ENDPOINT, AZURE_SEARCH_KEY, AZURE_SEARCH_ENDPOINT are all set.'
    );
}

// Create AI components
const model = new OpenAIModel({
    // Azure OpenAI Support
    azureApiKey: process.env.AZURE_OPENAI_KEY!,
    azureDefaultDeployment: deployment,
    azureEndpoint: process.env.AZURE_OPENAI_ENDPOINT!,
    azureApiVersion: '2024-02-15-preview',

    // Request logging
    logRequests: true
});

const prompts = new PromptManager({
    promptsFolder: path.join(__dirname, '../src/prompts')
});

const planner = new ActionPlanner({
    model,
    prompts,
    defaultPrompt: async () => {
        const template = await prompts.getPrompt('chat');

        template.config.completion.model = deployment;
        (template.config.completion as any).data_sources = [
            {
                type: 'azure_search',
                parameters: {
                    endpoint: process.env.AZURE_SEARCH_ENDPOINT,
                    index_name: process.env.AZURE_SEARCH_INDEX || 'healthplan',
                    authentication: {
                        type: 'api_key',
                        key: process.env.AZURE_SEARCH_KEY
                    }
                }
            }
        ];

        return template;
    }
});

// Define storage and application
const storage = new MemoryStorage();
export const app = new Application<ApplicationTurnState>({
    storage: storage,
    ai: { planner },
    adapter: new TeamsAdapter(
        {},
        new ConfigurationServiceClientCredentialFactory({
            MicrosoftAppId: process.env.BOT_ID, // Set to "" if using the Teams Test Tool
            MicrosoftAppPassword: process.env.BOT_PASSWORD, // Set to "" if using the Teams Test Tool
            MicrosoftAppType: 'MultiTenant'
        })
    )
});

app.error(async (context: TurnContext, err: any) => {
    // This check writes out errors to console log .vs. app insights.
    // NOTE: In production environment, you should consider logging this to Azure
    //       application insights.
    error(`[onTurnError] unhandled error: ${err}`);
    error(err);

    // Send a trace activity, which will be displayed in Bot Framework Emulator
    await context.sendTraceActivity(
        'OnTurnError Trace',
        `${err}`,
        'https://www.botframework.com/schemas/error',
        'TurnError'
    );

    // Send a message to the user
    await context.sendActivity('The bot encountered an error or bug.');
    await context.sendActivity('To continue to run this bot, please fix the bot source code.');
});

app.ai.action(
    AI.SayCommandActionName,
    async (context: TurnContext, state: ApplicationTurnState, data: { response: Message<string> }) => {
        const attachment = CardFactory.adaptiveCard(createResponseCard(data.response));
        const activity = MessageFactory.attachment(attachment);
        await context.sendActivity(activity);
        return '';
    }
);
