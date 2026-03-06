from translator import OfflineTranslator

t = OfflineTranslator()
t.install_argos_model_online("en", "zh")
print("Argos en->zh 模型安装完成")
