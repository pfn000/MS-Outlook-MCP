FROM python:3.11-slim

WORKDIR /app

# Copy requirements and install dependencies
COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

# Copy MCP server
COPY microsoft_mcp.py .

# Create non-root user for security
RUN useradd -m -u 1000 mcp_user && chown -R mcp_user:mcp_user /app
USER mcp_user

# Expose port for HTTP transport (optional)
EXPOSE 8000

# Run the MCP server
# Can be overridden with docker run flags
CMD ["python", "microsoft_mcp.py"]
