<p align="center"><strong>Lexicon</strong> is a general purpose agent that runs locally on your computer. It is based on OpenAI Codex CLI but is tuned for more general use cases.
</p>

---

## Quickstart

### Installing and running Lexicon
You currently need to build this yourself.

Then simply run `lexicon` to get started:

```shell
lexicon
```

### Using Lexicon with your ChatGPT plan

<p align="center">
  <img src="./.github/codex-cli-login.png" alt="Lexicon login" width="80%" />
  </p>

Run `lexicon` and select **Sign in with ChatGPT**. We recommend signing into your ChatGPT account to use Lexicon as part of your Plus, Pro, Team, Edu, or Enterprise plan. [Learn more about what's included in your ChatGPT plan](https://help.openai.com/en/articles/11369540-codex-in-chatgpt).

You can also use Lexicon with an API key, but this requires [additional setup](./docs/authentication.md#usage-based-billing-alternative-use-an-openai-api-key). If you previously used an API key for usage-based billing, see the [migration steps](./docs/authentication.md#migrating-from-usage-based-billing-api-key).

### Model Context Protocol (MCP)

Lexicon can access MCP servers. To configure them, refer to the [config docs](./docs/config.md#mcp_servers).

### Configuration

Lexicon supports a rich set of configuration options, with preferences stored in `~/.lexicon/config.toml`. For full configuration options, see [Configuration](./docs/config.md).

### Execpolicy

See the [Execpolicy quickstart](./docs/execpolicy.md) to set up rules that govern what commands Lexicon can execute.

---

## Attribution

Lexicon is a fork of [OpenAI Codex CLI](https://github.com/openai/codex), licensed under Apache 2.0.

## License

This repository is licensed under the [Apache-2.0 License](LICENSE).
