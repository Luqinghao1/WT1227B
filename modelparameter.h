/*
 * 文件名: modelparameter.h
 * 文件作用: 项目参数单例类头文件
 * 功能描述:
 * 1. 管理项目核心数据和文件交互。
 * 2. [新增逻辑] savePlottingData 现在会将绘图数据写入单独的 "项目名_chart.json" 文件中。
 * 3. [新增逻辑] loadProject 会自动尝试读取同目录下的 "项目名_chart.json" 文件以恢复图表数据。
 */

#ifndef MODELPARAMETER_H
#define MODELPARAMETER_H

#include <QString>
#include <QObject>
#include <QJsonObject>
#include <QJsonDocument>
#include <QJsonArray>
#include <QMutex>

class ModelParameter : public QObject
{
    Q_OBJECT

public:
    static ModelParameter* instance();

    // ========================================================================
    // 项目文件管理
    // ========================================================================

    // 加载项目文件 (.pwt)
    // 同时会自动寻找并加载同名的 "_chart.json" 图表数据文件
    bool loadProject(const QString& filePath);

    // 保存基础参数到 .pwt 文件
    bool saveProject();

    // 关闭项目
    void closeProject();

    QString getProjectFilePath() const { return m_projectFilePath; }
    QString getProjectPath() const { return m_projectPath; }
    bool hasLoadedProject() const { return m_hasLoaded; }

    // ========================================================================
    // 数据存取
    // ========================================================================

    void setParameters(double phi, double h, double mu, double B, double Ct, double q, double rw, const QString& path);

    double getPhi() const { return m_phi; }
    double getH() const { return m_h; }
    double getMu() const { return m_mu; }
    double getB() const { return m_B; }
    double getCt() const { return m_Ct; }
    double getQ() const { return m_q; }
    double getRw() const { return m_rw; }

    // 保存拟合结果 (写入 .pwt)
    void saveFittingResult(const QJsonObject& fittingData);
    QJsonObject getFittingResult() const;

    // [修改] 保存绘图数据
    // 将 plots 数据保存到单独的 "[项目名]_chart.json" 文件中
    void savePlottingData(const QJsonArray& plots);

    // 获取绘图数据 (从内存缓存中读取，由 loadProject 预先加载)
    QJsonArray getPlottingData() const;

private:
    explicit ModelParameter(QObject* parent = nullptr);
    static ModelParameter* m_instance;

    bool m_hasLoaded;
    QString m_projectPath;
    QString m_projectFilePath;

    // 缓存完整的JSON对象
    QJsonObject m_fullProjectData;

    // 基础参数
    double m_phi;
    double m_h;
    double m_mu;
    double m_B;
    double m_Ct;
    double m_q;
    double m_rw;

    // 辅助：获取图表数据文件的路径
    // 例如: D:/proj/demo.pwt -> D:/proj/demo_chart.json
    QString getPlottingDataFilePath() const;
};

#endif // MODELPARAMETER_H
