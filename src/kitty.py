# Keypirinha | A semantic launcher for Windows | http://keypirinha.com

import keypirinha as kp
import keypirinha_util as kpu
import os
import os.path
import winreg
import urllib.parse

class KiTTY(kp.Plugin):
    """
    Launch KiTTY sessions.

    This plugin automatically detects the installed version of the official
    KiTTY distribution and lists its configured sessions so they can be launched
    directly without having to pass through the sessions selection dialog.
    """

    DIST_SECTION_PREFIX = "dist/" # lower case
    EXE_NAME_OFFICIAL = "KITTY.EXE"

    default_icon_handle = None
    distros = {}

    def __init__(self):
        super().__init__()
        self._debug = True

    def on_start(self):
        self._read_config()

    def on_catalog(self):
        self._read_config()

        catalog = []
        for distro_name, distro in self.distros.items():
            if not distro['enabled']:
                continue
            # catalog the executable
            catalog.append(self.create_item(
                category=kp.ItemCategory.FILE,
                label=distro['label'],
                short_desc="",
                target=distro['exe_file'],
                args_hint=kp.ItemArgsHint.ACCEPTED,
                hit_hint=kp.ItemHitHint.KEEPALL,
                data_bag= kpu.kwargs_encode(
                    distro_name=distro_name
                ))
            )
        self.set_catalog(catalog)


    def on_suggest(self, user_input, items_chain):
        if not items_chain or items_chain[0].category() != kp.ItemCategory.FILE:
            return

        suggestions = []

        data_bag = kpu.kwargs_decode(items_chain[0].data_bag())
        sessions = self.distros[data_bag['distro_name']]['sessions']

        for session_name in sessions:
            if not user_input or kpu.fuzzy_score(user_input, session_name) > 0:
                suggestions.append(self.create_item(
                    category=kp.ItemCategory.REFERENCE,
                    label="{}".format(session_name),
                    short_desc='Launch "{}" session'.format(session_name),
                    target=kpu.kwargs_encode(
                        dist=data_bag['distro_name'], session=session_name),
                    args_hint=kp.ItemArgsHint.FORBIDDEN,
                    hit_hint=kp.ItemHitHint.IGNORE))

        self.set_suggestions(suggestions, kp.Match.ANY, kp.Sort.NONE)


    def on_execute(self, item, action):
        if item.category() == kp.ItemCategory.FILE:
            kpu.execute_default_action(self, item, action)
            return

        if item.category() != kp.ItemCategory.REFERENCE:
            return

        # extract info from item's target property
        try:
            item_target = kpu.kwargs_decode(item.target())
            distro_name = item_target['dist']
            session_name = item_target['session']
        except Exception as e:
            self.dbg(e)
            return

        # check if the desired distro is available and enabled
        if distro_name not in self.distros:
            self.warn('Could not execute item "{}". Distro "{}" not found.'.format(item.label(), distro_name))
            return
        distro = self.distros[distro_name]
        if not distro['enabled']:
            self.warn('Could not execute item "{}". Distro "{}" is disabled.'.format(item.label(), distro_name))
            return

        # check if the desired session still exists
        if session_name not in distro['sessions']:
            self.warn('Could not execute item "{}". Session "{}" not found in distro "{}".'.format(item.label(), session_name, distro_name))
            return

        # find the placeholder of the session name in the args list and execute
        sidx = distro['cmd_args'].index('%1')
        kpu.shell_execute(
            distro['exe_file'],
            args=distro['cmd_args'][0:sidx] + [session_name] + distro['cmd_args'][sidx+1:])


    def on_events(self, flags):
        if flags & kp.Events.PACKCONFIG:
            self.info("Configuration changed, rebuilding catalog...")
            self.on_catalog()


    def _read_config(self):
        if self.default_icon_handle:
            self.default_icon_handle.free()
            self.default_icon_handle = None
        self.distros = {}

        settings = self.load_settings()
        for section_name in settings.sections():
            if not section_name.lower().startswith(self.DIST_SECTION_PREFIX):
                continue

            dist_name = section_name[len(self.DIST_SECTION_PREFIX):]

            detect_method = getattr(self, "_detect_distro_{}".format(dist_name.lower()), None)
            if not detect_method:
                self.err("Unknown KiTTY distribution name: ", dist_name)
                continue

            dist_path = settings.get_stripped("path", section_name)
            dist_enable = settings.get_bool("enable", section_name)
            dist_file_based = settings.get_bool("file_based", section_name)

            dist_props = detect_method(
                dist_enable,
                settings.get_stripped("label", section_name),
                dist_path,
                dist_file_based)

            if not dist_props:
                if dist_path:
                    self.warn('KiTTY distribution "{}" not found in: {}'.format(dist_name, dist_path))
                elif dist_enable:
                    self.warn('KiTTY distribution "{}" not found'.format(dist_name))
                continue

            self.distros[dist_name.lower()] = {
                'orig_name': dist_name,
                'enabled': dist_props['enabled'],
                'label': dist_props['label'],
                'exe_file': dist_props['exe_file'],
                'cmd_args': dist_props['cmd_args'],
                'file_based': dist_file_based,
                'sessions': dist_props['sessions']}

            if dist_props['enabled'] and not self.default_icon_handle:
                self.default_icon_handle = self.load_icon(
                    "@{},0".format(dist_props['exe_file']))
                if self.default_icon_handle:
                    self.set_default_icon(self.default_icon_handle)


    def _detect_distro_official(self, given_enabled, given_label, given_path, given_file_based):
        dist_props = {
            'enabled': given_enabled,
            'label': given_label,
            'exe_file': None,
            'cmd_args': ['-load', '%1'],
            'file_based': given_file_based,
            'sessions': []}

        # label
        if not dist_props['label']:
            dist_props['label'] = "KiTTY"

        # enabled? don't go further if not
        if dist_props['enabled'] is None:
            dist_props['enabled'] = True
        if not dist_props['enabled']:
            return dist_props

        if dist_props['file_based'] is None:
            dist_props['file_based'] = False

        # find executable
        exe_file = None
        if given_path:
            exe_file = os.path.join(given_path, self.EXE_NAME_OFFICIAL)
            if not os.path.exists(exe_file):
                exe_file = None
        if not exe_file:
            exe_file = self._autodetect_startmenu(self.EXE_NAME_OFFICIAL, "KiTTY.lnk")
        if not exe_file:
            exe_file = self._autodetect_official_progfiles()
        if not exe_file:
            exe_file = self._autodetect_path(self.EXE_NAME_OFFICIAL)
        #if not exe_file:
        #    exe_file = self._autodetect_startmenu(self.EXE_NAME_OFFICIAL, "*kitty*.lnk")
        if not exe_file:
            return None
        dist_props['exe_file'] = exe_file

        # list configured sessions
        if not dist_props['file_based']:
            try:
                hkey = winreg.OpenKey(
                    winreg.HKEY_CURRENT_USER,
                    'Software\\9bis.com\\KiTTY\\Sessions',
                    access=winreg.KEY_READ | winreg.KEY_ENUMERATE_SUB_KEYS)
                index = 0
                while True:
                    try:
                        dist_props['sessions'].append(urllib.parse.unquote(
                            winreg.EnumKey(hkey, index), encoding='mbcs'))
                        index += 1
                    except OSError:
                        break
                winreg.CloseKey(hkey)

            except OSError:
                pass
        else:
            exe_directory = os.path.dirname(os.path.abspath(exe_file))
            dist_props['sessions'] = self._get_sessions_from_folder(exe_directory)

        return dist_props


    def _autodetect_official_progfiles(self):
        for hive in ('%PROGRAMFILES%', '%PROGRAMFILES(X86)%'):
            exe_file = os.path.join(
                os.path.expandvars(hive), "KiTTY", self.EXE_NAME_OFFICIAL)
            if os.path.exists(exe_file):
                return exe_file


    def _autodetect_startmenu(self, exe_name, name_pattern):
        known_folders = (
            "{625b53c3-ab48-4ec1-ba1f-a1ef4146fc19}", # FOLDERID_StartMenu
            "{a4115719-d62e-491d-aa7c-e74b8be3b067}") # FOLDERID_CommonStartMenu

        found_link_files = []
        for kf_guid in known_folders:
            try:
                known_dir = kpu.shell_known_folder_path(kf_guid)
                found_link_files += [
                    os.path.join(known_dir, f)
                    for f in kpu.scan_directory(
                        known_dir, name_pattern, kpu.ScanFlags.FILES, -1)]
            except Exception as e:
                self.dbg(e)
                pass

        for link_file in found_link_files:
            try:
                link_props = kpu.read_link(link_file)
                if (link_props['target'].lower().endswith(exe_name) and
                        os.path.exists(link_props['target'])):
                    return link_props['target']
            except Exception as e:
                self.dbg(e)
                pass

        return None


    def _autodetect_path(self, exe_name):
        path_dirs = [
            os.path.expandvars(p.strip())
                for p in os.getenv("PATH", "").split(";") if p.strip() ]

        for path_dir in path_dirs:
            exe_file = os.path.join(path_dir, exe_name)
            if os.path.exists(exe_file):
                return exe_file

        return None


    def _get_sessions_from_folder(self, exe_folder):
        sessions = []
        for file in os.listdir(os.path.join(exe_folder, 'Sessions')):
            session_file_name = urllib.parse.unquote(file)
            sessions.append(session_file_name)

        return sessions