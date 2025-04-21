def validate_factory_split(self, cif_invoice_path, import_invoice_dir):
        """校验工厂拆分是否正确
        
        Args:
            cif_invoice_path: CIF发票路径
            import_invoice_dir: 进口发票目录
            
        Returns:
            tuple: (bool, str) 是否通过校验，错误信息
        """
        try:
            # 读取CIF发票，查找工厂列
            logging.info(f"读取CIF发票: {cif_invoice_path}")
            cif_df = pd.read_excel(cif_invoice_path)
            logging.info(f"CIF发票列名: {list(cif_df.columns)}")
            
            # 尝试识别工厂列
            factory_col = None
            # 尝试常用的工厂列名模式
            factory_patterns = ["Plant Location", "工厂地点", "工厂", "factory", "Factory", "FACTORY", 
                                "Plant", "plant", "Location", "location", "工厂名称", "Supplier"]
            
            for pattern in factory_patterns:
                for col in cif_df.columns:
                    if pattern in str(col):
                        factory_col = col
                        logging.info(f"找到工厂列: {col}")
                        break
                if factory_col:
                    break
                    
            # 如果直接匹配没找到，尝试更广泛的匹配
            if not factory_col:
                for col in cif_df.columns:
                    col_str = str(col).lower()
                    if 'factory' in col_str or '工厂' in col_str or 'plant' in col_str or 'location' in col_str or 'supplier' in col_str:
                        factory_col = col
                        logging.info(f"通过部分匹配找到工厂列: {col}")
                        break
            
            # 检查进口发票目录是否存在
            if not os.path.exists(import_invoice_dir):
                return False, f"进口发票目录不存在: {import_invoice_dir}"
                
            # 获取目录下所有Excel文件
            import_files = [f for f in os.listdir(import_invoice_dir) 
                            if f.endswith('.xlsx') or f.endswith('.xls')]
            
            if not import_files:
                return False, f"进口发票目录中没有找到Excel文件: {import_invoice_dir}"
                
            logging.info(f"进口发票目录中的文件: {import_files}")
            
            # 检查是否有工厂拆分的标志
            has_factory_split = False
            for f in import_files:
                if 'reimport_' in f:
                    has_factory_split = True
                    break
            
            # 特殊情况：如果只有一个默认工厂文件，认为是合法的
            if len(import_files) == 1 and ('默认工厂' in import_files[0] or '默认工厂'.encode('utf-8').decode('latin1') in import_files[0]):
                return True, "使用了默认工厂，验证通过"
                
            # 如果没有工厂列但有多个进口文件，说明可能是手动拆分的
            if not factory_col and len(import_files) > 1:
                logging.warning("未找到工厂列，但有多个进口文件，可能是手动拆分")
                return True, "未找到工厂列但检测到多个进口文件，验证通过（可能是手动拆分）"
            
            # 如果找到工厂列，验证每个工厂是否都有对应的文件
            if factory_col:
                factories = cif_df[factory_col].dropna().unique()
                logging.info(f"CIF发票中的工厂: {factories}")
                
                # 如果没有工厂值但使用了默认工厂文件，认为是合法的
                if len(factories) == 0:
                    for f in import_files:
                        if '默认工厂' in f or '默认工厂'.encode('utf-8').decode('latin1') in f:
                            return True, "CIF中无工厂值，使用了默认工厂，验证通过"
                
                # 验证每个工厂是否都有对应的文件
                missing_factories = []
                for factory in factories:
                    factory_str = str(factory).replace('/', '_').replace('\\', '_')
                    found = False
                    
                    for f in import_files:
                        # 处理可能的编码问题
                        try:
                            if factory_str in f:
                                found = True
                                break
                        except:
                            # 尝试不同的编码处理
                            if factory_str.encode('utf-8').decode('latin1') in f:
                                found = True
                                break
                    
                    if not found:
                        missing_factories.append(factory)
                
                if missing_factories:
                    return False, f"以下工厂没有对应的进口发票文件: {missing_factories}"
                
                return True, "所有工厂都有对应的进口发票文件，验证通过"
            
            # 如果找不到工厂列但使用了默认工厂，也认为是合法的
            for f in import_files:
                if '默认工厂' in f or '默认工厂'.encode('utf-8').decode('latin1') in f:
                    return True, "未找到工厂列，使用了默认工厂，验证通过"
            
            return False, "未找到工厂列，且没有默认工厂的进口发票文件"
            
        except Exception as e:
            logging.error(f"验证工厂拆分时出错: {str(e)}")
            import traceback
            traceback.print_exc()
            return False, f"验证失败: {str(e)}" 