"""!
@brief Office Janitor package root.
@details Modules under this namespace coordinate detection, planning,
uninstallation, and cleanup tasks for Microsoft Office installations per the
project specification.
"""

__all__ = [
    "main",
    "detect",
    "plan",
    "scrub",
    "msi_uninstall",
    "c2r_uninstall",
    "c2r_odt",
    "c2r_integrator",
    "licensing",
    "registry_tools",
    "fs_tools",
    "processes",
    "tasks_services",
    "logging_ext",
    "command_runner",
    "exec_utils",
    "restore_point",
    "ui",
    "tui",
    "constants",
    "safety",
    "confirm",
    "version",
    "guid_utils",
    "msi_components",
    "odt_build",
    "repair",
    "auto_repair",
    "cli_help",
]
