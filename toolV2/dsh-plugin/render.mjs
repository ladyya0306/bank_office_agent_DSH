export function renderOfficeOutput(value) {
  if (!value || typeof value !== 'object' || !value.delivery || typeof value.delivery !== 'object') {
    return JSON.stringify(value);
  }
  const { attachments: _attachments, attachment_groups, ...delivery } = value.delivery;
  const versionByPath = new Map((Array.isArray(delivery.versions) ? delivery.versions : [])
    .filter((version) => version && typeof version.path === 'string')
    .map((version) => [version.path.replace(/\\/g, '/').toLocaleLowerCase(), version]));
  const present_calls = Array.isArray(attachment_groups)
    ? attachment_groups.map((group) => ({
      files: (Array.isArray(group) ? group : []).map((file) => {
        const filePath = String(file);
        const outputName = filePath.split(/[\\/]/).pop();
        const version = versionByPath.get(filePath.replace(/\\/g, '/').toLocaleLowerCase());
        const templateName = version?.template_path?.split(/[\\/]/).pop()
          || version?.template?.split(/[\\/]/).pop();
        const parentName = filePath.split(/[\\/]/).slice(-2, -1)[0];
        const description = templateName
          ? (parentName ? `${parentName} / ${templateName}` : templateName)
          : outputName;
        return { path: filePath, description };
      }),
    }))
    : [];
  return JSON.stringify({ ...value, delivery: { ...delivery, present_calls } });
}
