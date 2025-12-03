from PyInstaller.utils.hooks import collect_submodules, get_module_file_attribute

# 确保 numpy 被正确收集
hiddenimports = collect_submodules('numpy')

# 移除可能导致问题的子模块
problematic = ['numpy.core._multiarray_tests', 'numpy.testing']
hiddenimports = [m for m in hiddenimports if not any(p in m for p in problematic)]
