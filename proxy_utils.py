import zipfile
import re
import os


class ProxyUtil:
    @classmethod
    def delete_work_file(cls, ext_zip_file_path):
        if os.path.exists(ext_zip_file_path):
            os.remove(ext_zip_file_path)

    @classmethod
    def create_proxy_extentions(cls, proxy_url: str, tmp_work_dir: str = ""):
        proxy_pattern = re.compile(
            r"(?P<protocol>https?)://(?P<username>[^:]+):(?P<password>[^@]+)@(?P<host>[^:]+):(?P<port>\d+)"
        )
        match = proxy_pattern.match(proxy_url)
        if match:
            proxy_host = match.group("host")
            proxy_port = match.group("port")
            proxy_username = match.group("username")
            proxy_password = match.group("password")

        manifest_json = """
        {
            "version": "1.0.0",
            "manifest_version": 2,
            "name": "Edge Proxy",
            "permissions": [
                "proxy",
                "tabs",
                "unlimitedStorage",
                "storage",
                "<all_urls>",
                "webRequest",
                "webRequestBlocking"
            ],
            "background": {
                "scripts": ["background.js"]
            },
            "minimum_edge_version":"79.0.0"
        }
        """

        background_js = f"""
        var config = {{
                mode: "fixed_servers",
                rules: {{
                singleProxy: {{
                    scheme: "http",
                    host: "{proxy_host}",
                    port: parseInt({proxy_port})
                }},
                bypassList: ["localhost"]
                }}
            }};
        
        chrome.proxy.settings.set({{value: config, scope: "regular"}}, function() {{}});
        
        function callbackFn(details) {{
            console.log('Authenticating with proxy:', '{proxy_username}', '{proxy_password}');
            return {{
                authCredentials: {{
                    username: "{proxy_username}",
                    password: "{proxy_password}"
                }}
            }};
        }}
        
        chrome.webRequest.onAuthRequired.addListener(
            callbackFn,
            {{urls: ["<all_urls>"]}},
            ['blocking']
        );
        """

        pluginfile = os.path.join(tmp_work_dir, "proxy_auth_plugin.zip")

        with zipfile.ZipFile(pluginfile, "w") as zp:
            zp.writestr("manifest.json", manifest_json)
            zp.writestr("background.js", background_js)

        return pluginfile
