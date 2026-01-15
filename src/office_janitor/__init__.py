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
    "registry_office",
    "registry_wi_cleanup",
    "registry_user",
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
    "repair_odt",
    "auto_repair",
    "cli_help",
]
