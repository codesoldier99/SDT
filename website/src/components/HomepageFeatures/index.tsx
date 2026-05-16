import type {ReactNode} from 'react';
import clsx from 'clsx';
import Heading from '@theme/Heading';
import styles from './styles.module.css';

type FeatureItem = {
  title: string;
  Svg: React.ComponentType<React.ComponentProps<'svg'>>;
  description: ReactNode;
};

const FeatureList: FeatureItem[] = [
  {
    title: '官方计划驱动',
    Svg: require('@site/static/img/undraw_docusaurus_mountain.svg').default,
    description: (
      <>
        所有网站栏目、讲义模板和实验样板均以升级版授课计划 PDF 为唯一核心依据。
      </>
    ),
  },
  {
    title: '工程化课程组织',
    Svg: require('@site/static/img/undraw_docusaurus_tree.svg').default,
    description: (
      <>
        课程围绕指标分析、驱动开发、GNSS/RTK、IMU 标定、ESKF 与 AI 增强感知展开。
      </>
    ),
  },
  {
    title: '年度版本化更新',
    Svg: require('@site/static/img/undraw_docusaurus_react.svg').default,
    description: (
      <>
        使用 Docusaurus 与 GitHub Pages 发布课程网站，并为后续学年保留版本归档能力。
      </>
    ),
  },
];

function Feature({title, Svg, description}: FeatureItem) {
  return (
    <div className={clsx('col col--4')}>
      <div className={clsx('text--center', styles.featureVisual)}>
        <Svg className={styles.featureSvg} role="img" />
      </div>
      <div className={clsx('text--center padding-horiz--md', styles.featureCard)}>
        <Heading as="h3">{title}</Heading>
        <p>{description}</p>
      </div>
    </div>
  );
}

export default function HomepageFeatures(): ReactNode {
  return (
    <section className={styles.features}>
      <div className="container">
        <div className="row">
          {FeatureList.map((props, idx) => (
            <Feature key={idx} {...props} />
          ))}
        </div>
      </div>
    </section>
  );
}
