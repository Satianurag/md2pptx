# Font Demo: Inter & JetBrains Mono

Technical Presentation for Font Testing

## Executive Summary

This presentation demonstrates the dual-font system using **Inter** for all body text and UI elements, while **JetBrains Mono** handles code blocks, YAML configurations, and technical annotations.

## The Tech Stack

Our system leverages modern typography principles:

- Clean sans-serif fonts for readability
- Monospace fonts for technical precision
- Consistent hierarchy across all slides

## Configuration Example

```yaml
system:
  name: "GhostYield"
  version: "2.1.0"
  fonts:
    primary: "Inter"
    code: "JetBrains Mono"
  
blockchain:
  protocol: "Charms"
  network: "Base Sepolia"
  zk_prover: "Circom & SnarkJS"
```

## API Integration

The following code demonstrates our API structure:

```javascript
const config = {
  api_key: "sk_live_...",
  endpoint: "https://api.ghostyield.io",
  timeout: 30000
};

async function fetchMetrics() {
  const response = await fetch(`${config.endpoint}/v1/metrics`);
  return response.json();
}
```

## Data Pipeline

Our data flow consists of three key stages:

1. **Ingestion**: Raw data capture from multiple sources
2. **Processing**: `transform()` and `validate()` operations
3. **Output**: Formatted results in JSON structure

## Key Metrics

| Metric | Value | Target |
|--------|-------|--------|
| Latency | 45ms | < 100ms |
| Throughput | 2,400 TPS | > 2,000 TPS |
| Uptime | 99.99% | 99.9% |

## Conclusion

The combination of `Inter` for prose and `JetBrains Mono` for technical content creates a professional, readable presentation that follows modern design standards.

Key takeaways:
- Use `variable_name` for inline code references
- Code blocks render with monospace font automatically
- All body text uses Inter for optimal readability
