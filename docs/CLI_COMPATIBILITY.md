# OffScrub CLI Compatibility Matrix

| Legacy Script | Switch | Native Equivalent | Notes |
| --- | --- | --- | --- |
| All | `/PREVIEW`, `/DETECTONLY` | `--dry-run` | Runs detection/cleanup without executing uninstallers. |
| All | `/QUIET`, `/PASSIVE` | `--quiet` | Suppresses human logger INFO chatter; machine logs remain. |
| All | `/NOREBOOT` | `--no-reboot` | Suppresses reboot recommendations bit; exit code keeps reboot bit otherwise. |
| All | `/TR`, `/TESTRERUN` | rerun count | Executes uninstall twice to mirror OffScrub test reruns. |
| All | `/S`, `/SKIPSD` | `skip_shortcut_detection` | Skips Start Menu shortcut cleanup. |
| All | `/KL`, `/KEEPLICENSE` | `keep_license` | Skips license/cache cleanup when set. |
| C2R | `/OFFLINE`, `/FORCEOFFLINE` | `offline` flag | Marks Click-to-Run target offline for uninstall. |
| C2R | `/LOG <path>` | `log_directory` | Path captured for parity; native logging uses `logging_ext`. |
| C2R | `/RETERRORORSUCCESS` | return-success-on-error | Normalises non-zero exit codes to success (reboot bit preserved). |
| C2R | `/FORCEARPUNINSTALL` | `force` flag | Honoured as data flag; native uninstall already attempts setup fallback. |
| MSI | `/DELETEUSERSETTINGS` | `delete_user_settings` | Removes user settings/templates when requested. |
| MSI | `/KEEPUSERSETTINGS` | `keep_user_settings` | Skips user settings cleanup. |
| MSI | `/CLEARADDINREG` | `clear_addin_registry` | Purges add-in registry keys. |
| MSI | `/REMOVEVBA` | `remove_vba` | Removes VBA registry keys and caches. |
| MSI | `/OSE` | `ose` flag | Captured for parity; OSE-specific operations are guarded in scrub flows. |
| MSI | `/FASTREMOVE`, `/BYPASS`, `/SCANCOMPONENTS`, `/REMOVEOSPP` | data flags | Captured for compatibility; mapped into product dicts for downstream handling. |

Additional notes:
- Stage logs are emitted to human and machine loggers (`stage0_detection`, `stage1_uninstall`) to mirror OffScrub sequencing.
- Return code bitmask preserves reboot recommendation in bit `2`; `/RETERRORORSUCCESS` lowers other failures to success for batch compatibility.
